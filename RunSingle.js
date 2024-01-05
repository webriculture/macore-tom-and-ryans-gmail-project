const fs = require("fs");
const readline = require('readline');
const { google } = require('googleapis');
const getEmails = require('./Main.js');

// Load client secrets from a file (you need to create this file with your credentials)
const credentials = require('./credentials.json');

// Create an OAuth2 client
const { client_secret, client_id, redirect_uris } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);


if (fs.existsSync('tokens.json')){
    const storedTokens = fs.readFileSync('tokens.json');
    oAuth2Client.setCredentials(JSON.parse(storedTokens));
    getEmails(oAuth2Client);
} else {
const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: 'https://www.googleapis.com/auth/gmail.readonly', // Add necessary scopes
    response_type: 'code',
    prompt: 'consent',
  });
  
  console.log('Authorize this app by visiting this URL:', authUrl);
  //exec('start ' + authUrl);

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  
  rl.question('Enter the code from that page here: ', async (code) => {
    rl.close();
    const { tokens } = await oAuth2Client.getToken(code);
    console.log(tokens);
    //oAuth2Client.getToken(code, (err, tokens) => {
        //if (err) return console.error('Error retrieving access token:', err);
        oAuth2Client.setCredentials(tokens);
        fs.writeFileSync('tokens.json', JSON.stringify(tokens)); // Save the tokens
      //});
    //getEmails();
  });
}