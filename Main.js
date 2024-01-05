const { google } = require('googleapis');
const fs = require('fs');
const fsPromise = require('fs').promises;
const path = require('path');
const puppeteer = require('puppeteer');
const { PDFDocument } = require('pdf-lib');

const root = '//macorebx/CLIENT-FILES/'

let gmail;
let currentSubjectLine = "";
let shouldBreak = false;

async function getEmails(oAuth2Client) {
  gmail = google.gmail({ version: 'v1', auth: oAuth2Client });

  console.log("Now fetching emails");

  try {
    const res = await gmail.users.messages.list({
      userId: 'me', // 'me' refers to the authenticated user
      labelIds: ['INBOX'], // Search query to get inbox emails
      maxResults: 20,
    });
    const sentRes = await gmail.users.messages.list({
      userId: 'me', // 'me' refers to the authenticated user
      labelIds: ['SENT'], // Search query to get sent emails
      maxResults: 10,
    });

    if (res.data.messages) {
      for (const message of res.data.messages) {
        await processMessage(message);
        if (shouldBreak) { break; }
      }
    }

    console.log("Now checking sent");

    shouldBreak = false;

    if (sentRes.data.messages) {
      for (const message of sentRes.data.messages) {
        await processMessage(message);
        if (shouldBreak) { break; }
      }
    }
  }
  catch (error) {
    console.error('Error fetching emails:', error);
  }
}

async function processMessage(message) {
  const res = await gmail.users.messages.get({
    userId: 'me',
    id: message.id
  });

  const metadata = await getMetaData(res);

  try {
    const email = res.data;
    emailBody = await getEmailBody(email);

    let subjectLine = "";
    let subjectRaw = '';

    for (var headerIndex = 0; headerIndex < email.payload.headers.length; headerIndex++) {
      if (email.payload.headers[headerIndex].name == 'Subject') {
        subjectRaw = email.payload.headers[headerIndex].value;
        subjectLine = subjectRaw.replace(/[^a-z0-9#_-]/gi, ' ');
        currentSubjectLine = subjectLine;
      }
    }

    let savePath = await getPath(subjectRaw);

    /*
    if (!savePath) {
      savePath = await createFolder(subjectRaw);
    }
    */
    if (!savePath.startsWith("NONE")) {
      /*
      convertToPDF(emailBody).then((pdfBytes) => {
        fs.writeFileSync(savePath + subjectLine + '.pdf', pdfBytes);
        console.log('Email saved as ' + savePath + subjectLine + '.pdf');
      }).catch((error) => {
        console.error('Error converting to PDF:', error);
      });
      */

      let suffix = "";
      if (subjectRaw.includes("Macore.com - Order Received")) {
        const dateObject = new Date(email.internalDate);
        const formattedDate = dateObject.toLocaleString(
          'default',
          {
            year: 'numeric',
            month: 'numeric',
            day: 'numeric',
            hour: 'numeric',
            minute: 'numeric',
            second: 'numeric'
          }
        );

        // Example filename using the formatted date
        suffix = `${formattedDate.replace(/[^\w\s]/gi, '')}`;
      }
      else {
        let fromName = metadata.from.split(" <")[0];
        suffix = fromName.replace(".", "_").replace("@", "_at_");
      }

      const fullPath = savePath + subjectLine + " " + suffix + ".pdf";

      if (await fs.existsSync(fullPath)) {
        const keywords = await readMetadata(fullPath);
        if (keywords == email.id) {
          // shouldBreak = true;
        }
      }

      if (!shouldBreak) {
        const attachments = await readAttachments(email);
        await convertToPDF_withHeader(emailBody, metadata, fullPath, email.id, email.payload.parts);

        let attachSavePath = savePath;
        const csrFolder = "/Emails/CSR-Client/";

        if (savePath.includes(csrFolder)) {
          attachSavePath = savePath.substring(0, savePath.indexOf(csrFolder)) + "/Proofs to Client/";
        }

        downloadAttachments(email, attachSavePath);
        // Process the email body, embed images, convert to PDF, etc.
        // Continue with the processing logic similar to the previous example
      }
    }
  }
  catch (err) {
    logError(err, res);
  }
}

async function checkParts(parts){
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];

    // Check if the part contains the email body (assuming text/plain)
    if (part.mimeType === 'text/html' && part.body && part.body.data) {
      return Buffer.from(part.body.data, 'base64').toString('utf-8');
    }
    else if (part.parts) {
      return checkParts(part.parts);
    }
  }
}

function getEmailBody(message) {
  const parts = message.payload.parts;

  if (
      message.payload &&
      message.payload.body &&
      message.payload.body.data
  ) {
      const emailBody = Buffer.from(message.payload.body.data, 'base64').toString('utf-8');
      return emailBody;
  }
  else if (parts && parts.length > 0) {
    return checkParts(parts);
  }

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
  }

  // If no body part is found, handle accordingly
  return 'No body found in the email';
}

async function getPath(subjectLine){
  let path = 'NONE';
  subjectLine = subjectLine.toLowerCase();
  let orderFolder = await getOrderFolder(subjectLine);

  if (
      subjectLine.includes("inventory check") ||
      subjectLine.includes("date assignment") ||
      subjectLine.includes("outside service alert") ||
      subjectLine.includes("quote request")
  ) {
    path = orderFolder + "/Emails/OE Tech/";
  }
  else if (subjectLine.includes("order complete")) {
    path = orderFolder + "/Shipping/";
  }
  else if (subjectLine.includes("macore label order")) {
    path = orderFolder + "/Emails/CSR-Client/";
  }
  else if (subjectLine.includes('macore.com - order received')) {
    path = root + "/_NEW WEB ORDERS/";
  }
  else {
    path = orderFolder + "/";
  }

  return path;
}

async function getOrderFolder(subjectLine) {
  let match = subjectLine.match(/#(\d{4})\b/);
  let accountNumber = match ? match[0] : null;

  if (!accountNumber) { return "NONE"; }

  accountNumber = accountNumber.slice(1);

  let orderMatch = subjectLine.match(/#(\d{5})\b/);
  let orderNumber = orderMatch ? orderMatch[0] : null;

  if (!orderNumber) { return "NONE"; }

  orderNumber = orderNumber.slice(1);

  let accountFolder = await getAccountFolder(accountNumber);
  let orderFolder = await findSubfolder(root + accountFolder, orderNumber);

  return orderFolder;
}

async function getAccountFolder(accountNumber){
  const files = await fsPromise.readdir(root);
  let matchingFolder = files.find(file => {
    const folderPath = path.join(root, file);
    return fs.statSync(folderPath).isDirectory() && file.endsWith(`_${accountNumber}`);
  });

  if (!matchingFolder) {
    console.log("Creating account folder");
    matchingFolder = "_New Client _" + accountNumber;

    await fs.mkdirSync(root + "/" + matchingFolder);
    //await fs.cpSync(root + "/_NEW CLIENT FOLDER TEMPLATE/Client Name_9999", root + "/" + matchingFolder + "/", { recursive: true });
  }

  return matchingFolder;
}

async function findSubfolder(rootFolder, orderNumber) {
  const files = await fsPromise.readdir(rootFolder); // Read the contents of the root folder synchronously

  for (const file of files) {
    const folderPath = path.join(rootFolder, file);
    const stats = fs.statSync(folderPath);

    if (stats.isDirectory()) {
      const subFolderPath = path.join(folderPath, orderNumber);
      if (fs.existsSync(subFolderPath) && fs.statSync(subFolderPath).isDirectory()) {
        // Do something with the folder if needed
        return subFolderPath; // Return the path to the found folder
      }/* else {
        // If it's not the target folder, check its subfolders
        const subfolderPath = findSubfolder(folderPath);
        if (subfolderPath) {
          return subfolderPath; // Return the path if found in subfolder
        }
      }*/
    }
  }

  //Check if current year folder exists
  let yearFolder = rootFolder + "/" + new Date().getFullYear();
  if (await fs.existsSync(yearFolder) == false) {
    //Create if it doesn't
    await fs.mkdirSync(yearFolder);
  }

  //Create order folder within
  let subFolderPath = yearFolder + "/" + orderNumber;
  await fs.cpSync(root + "/_NEW ORDER FOLDER TEMPLATE/99999", subFolderPath, { recursive: true });
  return subFolderPath;
}

function downloadAttachments(email, path){
  const parts = email.payload.parts;
  if (parts) {
    parts.forEach(part => {
      if (part.filename) {
        const attachmentId = part.body.attachmentId;
        gmail.users.messages.attachments.get({
          userId: 'me',
          messageId: email.id,
          id: attachmentId,
        }, (err, attachment) => {
          if (err) return console.log('Error while fetching attachment:', err);

          const data = attachment.data;
          const fileData = Buffer.from(data.data, 'base64');
          fs.writeFileSync(path + part.filename, fileData);
          console.log('Attachment downloaded:', part.filename);
        });
      }
    });
  }
}

async function readAttachments(email) {
  const parts = email.payload.parts;
  let attachments = [];

  if (parts) {
    for (const part of parts) {
      if (part.filename) {
        console.log(part.filename);
        const attachmentId = part.body.attachmentId;
        const attachment = await gmail.users.messages.attachments.get({
          userId: 'me',
          messageId: email.id,
          id: attachmentId,
        });
        console.log("I might be reading an attachment named " + part.filename);
        if (attachment) {
          console.log("Reading attachment named " + part.filename);
          const data = attachment.data;
          const fileData = Buffer.from(data.data, 'base64');
          attachments.push(attachment);
        }
      }
    }
  }
  return attachments;
}

async function logError(err, subjectLine){
  const fileText = subjectLine + "\n" +"\n" + err;

  const formattedDate = new Date().toISOString().replace(/[:.]/g, '_');

  await fs.writeFileSync(path.join(__dirname, "/Logs/" + "Error_" + formattedDate + ".txt"), fileText);

  console.error(err);
}

async function convertToPDF_withHeader(emailBody, emailMetadata, savePath, id, parts) {
  const imageAttachments = parts.filter(part => part.mimeType.startsWith('image/'));
  //const imageAttachments = attachments.filter(attachment => attachment.type.startsWith('image/'));

  // Convert image attachments to base64 data URIs
  const imageURIs = await Promise.all(imageAttachments.map(async attachment => {
    console.log(attachment);
    const attachmentAsAttachment = await gmail.users.messages.attachments.get({
      userId: 'me',
      messageId: id,
      id: attachment.body.attachmentId,
    });
    const base64Data = attachmentAsAttachment.data.toString('base64');
    const mimeType = attachment.mimeType;
    return `data:${mimeType};base64,${base64Data}`;
  }));

  // Embed images in the HTML content
  let embeddedImagesHTML = '';
  imageURIs.forEach(uri => {
    embeddedImagesHTML += `<img src="${uri}" alt="Attachment Image"><br>`;
  });



  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();
  const to = escapePointyBrackets(emailMetadata.toName);

  // Create a header string from the email metadata
  const fullHTMLContent = `
<div class="email-header">
  <p><span style="font-size: 18px;"><b>${emailMetadata.subject}</b></span></p>
  <p><b>From: ${escapePointyBrackets(emailMetadata.fromName)}</b></p>
  <p><b>To: ${to}</b></p>
  <p><b>Date: ${emailMetadata.date}</b></p>
  <!-- Use <br> tags for manual line breaks if needed -->
</div>
<div class="email-body">
  ${emailBody}
  ${embeddedImagesHTML}
</div>
`;

  await page.setContent(fullHTMLContent);

  // Add CSS to style the header
  await page.addStyleTag({
    content: `
.email-header {
  margin: 5px 0; /* Adjust margin for spacing between lines */
  line-height: 0.1; /* Set the line-height to a smaller value (e.g., 1) */
  /* Add other styles to match Gmail header */
}
.email-body {
  /* Add styles for email body */
}
`,
  });

  const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });
  const pdfDoc = await PDFDocument.load(pdfBuffer);

  // Add custom metadata (hidden ID in this case)
  pdfDoc.setKeywords([id]); // Set keywords

  // Save the modified PDF
  const modifiedPdfBytes = await pdfDoc.save();

  if (savePath) {
    await fs.writeFileSync(savePath, modifiedPdfBytes);
    console.log(`Email saved as ${savePath}`);
  }

  await browser.close();
  return pdfBuffer;
}

async function getMetaData (res) {
  const headers = res.data.payload.headers;
  // Extract required metadata such as Subject, From, To, Date, etc.
  const metadata = {
    subject: '',
    from: '',
    fromName: '',
    to: '',
    toName: '',
    date: '',
    // Add other fields you need
  };

  // Loop through headers and extract required information
  headers.forEach((header) => {
    if (header.name === 'Subject') {
      metadata.subject = header.value;
    }
    else if (header.name === 'From') {
      metadata.fromName = header.value;
      const matches = header.value.match(/<([^>]+)>/); // Extract email address within angle brackets

      if (matches) {
        metadata.from = matches[1]; // Assign the extracted email address
      }
      else {
        // If no angle brackets found, extract email from the entire string
        const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/; // Simple email regex pattern
        const emailMatches = header.value.match(emailRegex);

        if (emailMatches) {
          metadata.from = emailMatches[0]; // Assign the extracted email address
        }
      }
    }
    else if (header.name === 'To') {
      metadata.toName = header.value;
      const matches = header.value.match(/<([^>]+)>/); // Extract email address within angle brackets

      if (matches) {
        metadata.to = matches[1]; // Assign the extracted email address
      }
      else {
        // If no angle brackets found, extract email from the entire string
        const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/; // Simple email regex pattern
        const emailMatches = header.value.match(emailRegex);

        if (emailMatches) {
          metadata.to = emailMatches[0]; // Assign the extracted email address
        }
      }
    }
    else if (header.name === 'Date') {
      metadata.date = header.value;
    }
    // Add other header fields you need
  });

  return metadata;
}

async function readMetadata(filePath) {
  const pdfBytes = await fs.readFileSync(filePath);
  const pdfDoc = await PDFDocument.load(pdfBytes);

  return pdfDoc.getKeywords();
}

function escapePointyBrackets(text){
  return text.replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

// let firstError = true;

process.on('uncaughtException', (error) => {
  logError(error, currentSubjectLine);

  /*
  if (firstError) {
    logError(error, currentSubjectLine);
    firstError = false;
  }
  else {
    throw(error);
  }
  */
});

module.exports = getEmails;
