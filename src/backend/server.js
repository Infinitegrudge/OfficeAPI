var express = require('express');
//import session from 'express-session';
let app = express();
const fs = require("fs");

app.use(express.static("./"));
app.use(express.json());
app.set('view engine', 'pug');

//const readline = require('readline-sync');

const settings = require('./appSettings');
const graphHelper = require('./graphHelper');

const PORT = 8000;

async function main() {
  console.log('JavaScript Graph Tutorial');

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

//   // Greet the user by name
//   await greetUserAsync();

//   const choices = [
//     'Display access token',
//     'List my inbox',
//     'Show drive',
//     'Make a Graph call'
//   ];

//   while (choice != -1) {
//     choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

//     switch (choice) {
//       case -1:
//         // Exit
//         console.log('Goodbye...');
//         break;
//       case 0:
//         // Display access token
//         await displayAccessTokenAsync();
//         break;
//       case 1:
//         // List emails from user's inbox
//         await listInboxAsync();
//         break;
//       case 2:
//         // Send an email message
//         await displayDriveAsync();
//         break;
//       case 3:
//         // Run any Graph code
//         await makeGraphCallAsync();
//         break;
//       default:
//         console.log('Invalid choice! Please try again.');
//     }
//   }
}

app.get(['/', '/home'], (req, res) => {
  fs.readFile("../frontend/index.html", function(err, data){
    if(err){
      res.status(500);
      res.end();
      return;
    }
    res.status(200);
    res.setHeader("Content-Type", "text/html");
    res.write(data);
    res.end();
  });
});

//main();
app.get("/login",  (req,res) => {
        initializeGraph(settings);
    }
)

function initializeGraph(settings) {
    graphHelper.initializeGraphForUserAuth(settings, (info) => {
        // Display the device code message to
        // the user. This tells them
        // where to go to sign in and provides the
        // code to use.
        console.log(info.message);
    });
}
  async function greetUserAsync() {
    try {
      const user = await graphHelper.getUserAsync();
      console.log(`Hello, ${user?.displayName}!`);
      // For Work/school accounts, email is in mail property
      // Personal accounts, email is in userPrincipalName
      console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);
    } catch (err) {
      console.log(`Error getting user: ${err}`);
    }
  }
  
  async function displayAccessTokenAsync() {
    try {
      const userToken = await graphHelper.getUserTokenAsync();
      console.log(`User token: ${userToken}`);
    } catch (err) {
      console.log(`Error getting user access token: ${err}`);
    }
  }

  app.get("/drive", async (req,res) => {
    if (req.session.loggedin) { 
      res.render("html/addWorkshop", {session: req.session});
    }
    else {
      console.log("not authenticated");
      res.status(401).json({'error': 'not authenticated'});
    }
  });

  async function displayDriveAsync() {
    try {
      const userDrive = await graphHelper.getDriveAsync();
      const drive = userDrive.value;
      console.log(drive);
    } catch (err) {
      console.log(`Error getting user drive: ${err}`);
    }
  }
  
  async function listInboxAsync() {
    try {
      const messagePage = await graphHelper.getInboxAsync();
      const messages = messagePage.value;
  
      // Output each message's details
      for (const message of messages) {
        console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
        console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
        console.log(`  Received: ${message.receivedDateTime}`);
      }
  
      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = messagePage['@odata.nextLink'] != undefined;
      console.log(`\nMore messages available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting user's inbox: ${err}`);
    }
  }


  const loadData = async () => {
	
    };
  loadData()
  .then(() => {

    app.listen(PORT);
    console.log("Listen on port:", PORT);

  })
  .catch(err => console.log(err));