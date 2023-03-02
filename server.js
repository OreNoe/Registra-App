const {google} = require('googleapis');
const admin = require('firebase-admin');

const serviceAccount = require('./serviceAccountKey.json');

admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
});

const db = admin.firestore();


const auth = new google.auth.GoogleAuth({
    keyFile: './serviceAccountKey.json',
    scopes: 'https://www.googleapis.com/auth/spreadsheets',
});

const sheets = google.sheets({version: 'v4', auth});

const sheetId = '1gwRs4BXhF2GN6-X0-34_mLUR8OjeGEeWSHvExclMH3Q';
const range = 'Sheet1';

db.collection('users').onSnapshot((snapshot) => {
    const rows = snapshot.docs.map((doc) => {
        const data = doc.data();
        console.log(data);
        return [data.name, data.surname ,data.email, data.lvl, data.state];
    });

    rows.unshift(['Name', 'Surname', 'Email', 'Nivel', 'Status']);

    snapshot.docChanges().forEach((change) => {
        if (change.type === 'removed') {

          sheets.spreadsheets.values.clear({
              spreadsheetId: sheetId,
              range: 'Sheet1!A2:E10000',
          }, (err, result) => {
              if (err) {
                  console.log(err);
              }else{
                  console.log('%d cells updated.', result.updatedCells);
              }
          });
        }else{
            sheets.spreadsheets.values.update({
              spreadsheetId: sheetId,
              range: range,
              valueInputOption: 'USER_ENTERED',
              resource: {
                  values: rows,
              },
          }, (err, result) => {
              if (err) {
                  console.log(err);
              }else
                  console.log('%d cells updated.', result.updatedCells);
          });
        }
    });
   
});
