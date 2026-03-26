const { google } = require('googleapis');
const fs = require('fs');
const fsPromises = require('fs').promises;
const path = require('path');
const puppeteer = require('puppeteer');
const { PDFDocument } = require('pdf-lib');

const ROOT_DIR = '//macorebx/CLIENT-FILES/'
const ORDER_LOG_PATH = path.join(__dirname, "/Logs/OrderNumberLog.txt");
const STANDARD_ORDER_TYPES = ["inventory check", "date assignment", "outside service alert", "quote request", "research request", "photo request", "labels for review", "coh change request", "revised order paperwork", "stringing pre-schedule", "call tag request", "samples request", "shipping/freight request", "check received", "credit memo", "watchlist", "credit to invoice client"]
const CSR_CLIENT_ORDER_TYPES = ["macore label order", "order status inquiry", "client order history"];
let gmail;
let currentSubjectLine = "";
let shouldBreak = false;
let startTime;
let fullSavedOrders;
let username;

//Main function to retrieve emails, called by RunAll and RunSingle
async function getEmails(oAuth2Client, canBreak) {
  startTime = new Date().toDateString() + ", " + new Date().toTimeString();

  fullSavedOrders = fs.existsSync(ORDER_LOG_PATH) ? await fsPromises.readFile(ORDER_LOG_PATH, 'utf8') : "";
  gmail = google.gmail({ version: 'v1', auth: oAuth2Client });

  try {
    const userInfo = await gmail.users.getProfile({userId: 'me',});
    username = userInfo.data.emailAddress.split("@")[0];
    console.log("Now fetching emails for " + username);

    await processEmailFolder('INBOX', canBreak, 500);
    await processEmailFolder('SENT', canBreak, 500);

    await fs.writeFileSync(ORDER_LOG_PATH, fullSavedOrders);

  } catch (err) {
    logError(err, currentSubjectLine);
  }
}

// Process a specific email folder (Inbox or Sent)
async function processEmailFolder(folder, canBreak, count) {
  console.log("Now checking " + folder);
  shouldBreak = false;
  const res = await gmail.users.messages.list({ userId: 'me', labelIds: [folder], maxResults: count });
  if (res.data.messages) {
    for (const message of res.data.messages) {
      await processMessage(message, canBreak, fullSavedOrders);
      if (shouldBreak) break;
    }
  }
}

//Process a specific email based on the id
async function processMessage(message, canBreak) {
  console.log("");
  console.log("");
  const res = await gmail.users.messages.get({userId: 'me', id: message.id});
  try {
    const metadata = await getEmailMetadata(res);
    const emailBody = await getEmailBody(res.data);
    const subjectRaw = await getHeaderValue(res.data.payload.headers);
    const subjectLine = await sanitizeSubject(subjectRaw);
    currentSubjectLine = subjectRaw;

    let savePath = await getPath(subjectRaw);

    if (!savePath.startsWith("NONE")) {
      const suffix = determineSuffix(subjectRaw, metadata, res.data);
      const fullPath = path.join(savePath, `${subjectLine} ${suffix}.pdf`);

      if (!await checkAndHandleExistingFile(fullPath, res.data, canBreak)) {
        await convertToPDF_withHeader(emailBody, metadata, fullPath, res.data.id, res.data.payload.parts);
        const attachSavePath = determineAttachmentPath(savePath);
        await downloadAttachments(res.data, attachSavePath);
      }
    }
  }
  catch (err) {
    logError(err, currentSubjectLine);
  }
}

//Scans the headers and finds the subject of the email
function getHeaderValue(headers) {
  return headers.find(header => header.name === 'Subject')?.value || '';
}

//Turns the subject line into a format compatible for file names
function sanitizeSubject(subject) {
  return subject.replace(/[^a-z0-9#_-]/gi, ' ');
}

//Determines the suffix to append after the subject line (before the file extension)
function determineSuffix(subjectRaw, metadata, email) {
  return subjectRaw.includes("Macore.com - Order Received") 
    ? email.internalDate 
    : metadata.from.split(" <")[0].replace(".", "_").replace("@", "_at_");;
}

//Checks if the file already exists, if it needs to be overwritten, and if we can cancel processing the remaining messages
async function checkAndHandleExistingFile(fullPath, email, canBreak) {
  if (await fs.existsSync(fullPath)) {
    console.log("File already exists at " + fullPath);
    const keywords = await getPDFMetadata(fullPath);

    if (keywords.includes(email.id) && canBreak) {
      console.log("ID's match. Moving to next array.");
      shouldBreak = true;
      return true;
    }

    if (keywords.includes(startTime)) {
      return true;
    }
  }
  return false;
}

//Determines if the attachments need to be saved at a different location from the email itself
function determineAttachmentPath(savePath) {
  const csrFolder = '\\Emails\\CSR-Client\\';
  return savePath.includes(csrFolder) ? path.join(savePath.substring(0, savePath.indexOf(csrFolder)), "Proofs to Client") : savePath;
}

//Checks each part of the email in order to locate the body. Used in case the body isn't found in the body
async function checkParts(parts){
  for (const part of parts) {
    // Check if the part contains the email body (assuming text/html)
    if (part.mimeType === 'text/html' && part.body?.data) {
      return Buffer.from(part.body.data, 'base64').toString('utf-8');
    }
    
    if (part.parts){
      return await checkParts(part.parts);
    }
  }
}

//Gets the body of the email
function getEmailBody(message) {  
    if (message.payload?.body?.data) {
        const emailBody = Buffer.from(message.payload.body.data, 'base64').toString('utf-8');
        return emailBody;
    } else if (message.payload.parts?.length > 0) {
      return checkParts(message.payload.parts);
    }

    // If no body part is found, handle accordingly
    return 'No body found in the email';
}

//Checks the subject line to determine the save path of the PDF
async function getPath(subjectLine){
    subjectLine = subjectLine.toLowerCase();
    const orderFolder = await findBaseFolder(subjectLine);
    console.log(subjectLine);

    if(STANDARD_ORDER_TYPES.some((v => subjectLine.includes(v)))){
      return path.join(orderFolder + "/Emails/OE Tech/");
    }
    else if (subjectLine.includes("order complete")){
      return path.join(orderFolder + "/Shipping/");
    }
    else if (CSR_CLIENT_ORDER_TYPES.some(v => subjectLine.includes(v)))
    {
      console.log("Sorting to CSR-Client");
      return path.join(orderFolder + "/Emails/CSR-Client/");
    }
    else if(subjectLine.includes('macore.com - order received')){
      return path.join(ROOT_DIR + "/_NEW ORDERS/Web Orders/");
    }
    else if (subjectLine.includes('new order -')){
      const newPath = path.join(ROOT_DIR, "_NEW ORDERS", username);
      await ensureDirectory(newPath);
      return newPath;
    }
    else{
      return path.join(orderFolder);
    }
}

//Makes sure a given directory exists
async function ensureDirectory(dirPath) {
  if (!await fs.existsSync(dirPath)) {
      await fs.mkdirSync(dirPath);
  }
}

//Gets the folder for the order based on a subject line
async function findBaseFolder(subjectLine) {
  const accountNumber = subjectLine.match(/#(\d{4})\b/)?.[0]?.slice(1);
  const orderNumber = subjectLine.match(/#(\d{5})\b/)?.[0]?.slice(1);

  if (!accountNumber || !orderNumber) return "NONE";

  if(!fullSavedOrders.includes(orderNumber)){
    fullSavedOrders += `#${orderNumber} for client #${accountNumber}\n`;
  }

  const accountFolder = await getAccountFolder(accountNumber);

  return await getOrderFolder(path.join(ROOT_DIR, accountFolder), orderNumber);
}

//Gets the folder for an account based on the account number
async function getAccountFolder(accountNumber){
      const files = await fsPromises.readdir(ROOT_DIR);
        let matchingFolder = files.find(file => file.endsWith(`_${accountNumber}`));

        if(!matchingFolder){
          console.log("Creating account folder");
          matchingFolder = `_New Client _${accountNumber}`;
    
          await fs.mkdirSync(path.join(ROOT_DIR, matchingFolder));
          //await fs.cpSync(root + "/_NEW CLIENT FOLDER TEMPLATE/Client Name_9999", root + "/" + matchingFolder + "/", { recursive: true });
        }

        return matchingFolder;
}

//Gets an order folder based off of the account folder and order number
async function getOrderFolder(accountFolder, orderNumber) {
      const files = await fsPromises.readdir(accountFolder); // Read the contents of the root folder synchronously
      
      for (const file of files) {
        const folderPath = path.join(accountFolder, file);
  
        if (await fs.statSync(folderPath).isDirectory()) {
          const subFolderPath = path.join(folderPath, orderNumber);
          if (await fs.existsSync(subFolderPath)) {
            // Do something with the folder if needed
            return subFolderPath; // Return the path to the found folder
          }
        }
      }

      return await createOrderFolder(accountFolder, orderNumber);
}

//Creates an order folder if it doesn't already exist
async function createOrderFolder(accountFolder, orderNumber){
  //Check if current year folder exists
  const yearFolder = path.join(accountFolder, new Date().getFullYear().toString());
  await ensureDirectory(yearFolder);
  
  //Create order folder within
  let subFolderPath = path.join(yearFolder, orderNumber);
  await fs.cpSync(path.join(ROOT_DIR, "/_NEW ORDER FOLDER TEMPLATE/99999"), subFolderPath, { recursive: true });
  return subFolderPath;
}

//Downloads the attachments of an email
async function downloadAttachments(email, savePath){
    const parts = email.payload.parts || [];
          for (const part of parts){
            if (part.filename) {
              const attachment = await gmail.users.messages.attachments.get({
                userId: 'me',
                messageId: email.id,
                id: part.body.attachmentId,
              });
                const fileData = Buffer.from(attachment.data.data, 'base64');
                fs.writeFileSync(path.join(savePath, part.filename), fileData);
                console.log('Attachment downloaded:', part.filename);
            }
          }

}

//Saves an error in the logs folder
async function logError(err, subjectLine){
  console.log(subjectLine);
  const errorMessage = `[Subject]: ${subjectLine}\n\n${err.stack || err}`;
  const formattedDate = new Date().toISOString().replace(/[:.]/g, '_');
  const errorFilePath = path.join(__dirname, "/Logs/" + "Error_" + formattedDate + ".txt");

  await fs.writeFileSync(errorFilePath, errorMessage);

  console.error(err);
}

//Converts the email's HTML to the bytes for a PDF
async function convertToPDF_withHeader(emailBody, emailMetadata, savePath, id, parts) {
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();

  emailBody = await convertCIDtoBase64(emailBody, id, parts);
  const fullHTMLContent = createHTMLContent(emailMetadata, emailBody);
  await page.setContent(fullHTMLContent);
  await addCSSStyle(page);

  const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });
  const pdfDoc = await PDFDocument.load(pdfBuffer);

    // Add custom metadata (hidden ID in this case)
    pdfDoc.setKeywords([id + ", " + startTime]); // Set keywords
    // Save the modified PDF
    const modifiedPdfBytes = await pdfDoc.save();

  if (savePath) {
    await fs.writeFileSync(savePath, modifiedPdfBytes);
    console.log(`Email saved as ${savePath}`);
  }

  await browser.close();
  return pdfBuffer;
}

function createHTMLContent(emailMetadata, emailBody) {
  const to = escapePointyBrackets(emailMetadata.toName);
  return `
  <div class="email-header">
  <p><span style="font-size: 18px;"><b>${emailMetadata.subject}</b></span></p>
  <p><b>From: ${escapePointyBrackets(emailMetadata.fromName)}</b></p>
  <p><b>To: ${to}</b></p>
  <p><b>Date: ${emailMetadata.date}</b></p>
  <!-- Use <br> tags for manual line breaks if needed -->
</div>
<div class="email-body">
  ${emailBody}
</div>
  `;
}

async function addCSSStyle(page) {
  const cssContent = `
    .email-header {
      margin: 5px 0;
      line-height: 0.1;
    }
    .email-body {
      /* Add styles for email body */
    }
  `;
  await page.addStyleTag({ content: cssContent });
}

//Extracts metadata from the email
async function getEmailMetadata (res){
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
      } else if (header.name === 'From') {
        metadata.fromName = header.value;
        const matches = header.value.match(/<([^>]+)>/); // Extract email address within angle brackets
        if (matches) {
          metadata.from = matches[1]; // Assign the extracted email address
        } else {
          // If no angle brackets found, extract email from the entire string
          const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/; // Simple email regex pattern
          const emailMatches = header.value.match(emailRegex);
          if (emailMatches) {
            metadata.from = emailMatches[0]; // Assign the extracted email address
          }
        }
      } else if (header.name === 'To') {
        metadata.toName = header.value;
        const matches = header.value.match(/<([^>]+)>/); // Extract email address within angle brackets
        if (matches) {
          metadata.to = matches[1]; // Assign the extracted email address
        } else {
          // If no angle brackets found, extract email from the entire string
          const emailRegex = /[\w.-]+@[\w.-]+\.[a-zA-Z]{2,}/; // Simple email regex pattern
          const emailMatches = header.value.match(emailRegex);
          if (emailMatches) {
            metadata.to = emailMatches[0]; // Assign the extracted email address
          }
        }
      } else if (header.name === 'Date') {
        metadata.date = header.value;
      }
      // Add other header fields you need
    });

    return metadata;
}

//Extracts metadata from a PDF
async function getPDFMetadata(filePath) {
    const pdfBytes = await fs.readFileSync(filePath);
    const pdfDoc = await PDFDocument.load(pdfBytes);
  
    return pdfDoc.getKeywords();
}

function escapePointyBrackets(text){
  return text.replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function getAttachmentId(parts, cid) {
  for (const part of flattenParts(parts)) {
    const contentId = part.headers?.['content-id'] || '';
    const xAttachmentId = part.headers?.['x-attachment-id'] || '';
    if (contentId.includes(cid) || xAttachmentId.includes(cid)) {
      return part;
    }
  }
  return null;
}

function flattenParts(parts) {
  let flatParts = [];
  for (const part of parts) {
    flatParts.push(part);
    if (part.parts) {
      flatParts = flatParts.concat(flattenParts(part.parts));
    }
  }
  return flatParts;
}

async function convertCIDtoBase64(input, id, parts) {
  const regex = /cid:([a-zA-Z0-9_]+)/g;
  const matches = input.match(regex); // Find matches synchronously

  if (matches) {
    for (const match of matches) {
        const cid = match.slice(4); // Extract CID value from match
        const part = getAttachmentId(parts, cid);
        if(part){
        const attachmentAsAttachment = await gmail.users.messages.attachments.get({
          userId: 'me',
          messageId: id,
          id: part.body.attachmentId,
        });
    
        const base64Data = attachmentAsAttachment.data.data.toString('base64').replaceAll("-", "+").replaceAll("_", "/");
        const mimeType = part.mimeType;
  
        
        // Replace each match asynchronously in the input string
        input = input.replace(match, `data:${mimeType};base64,${base64Data}`);
      }
    }
  }

  return input;
}

process.on('uncaughtException', (error) => {
  //if(firstError){
  logError(error, currentSubjectLine);
  /*firstError = false;
  }
  else{
    throw(error);
  }*/
});

module.exports = getEmails;
