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

db.collection('events').onSnapshot(snapshot => {

    snapshot.docChanges().forEach(change => {
        const doc = change.doc;
        const rows = snapshot.docs.map(doc => {
            const data = doc.data();
            console.log(data);
            return [data.name, data.surname ,data.email, data.lvl, data.encargado, data.state];
        });

        rows.unshift(['Name', 'Surname', 'Email', 'Nivel',  'Encargado', 'Status']);

        if (change.type == 'added' || change.type == 'modified'){
            const event = doc.data();
            const eventSheetName = event.name;
            const eventSheet = {
                spreadsheetId: sheetId,
                resource: {
                    requests: [{
                        addSheet: {
                            properties: {
                                title: eventSheetName
                            }
                        }
                    }]
                }
            };
            sheets.spreadsheets.batchUpdate(eventSheet, (err, res) => {
                if (err) return console.log('The API returned an error: ' + err);
                console.log('Sheet created');
            });

            db.collection('events').doc(doc.id).collection('users').onSnapshot(snapshot => {
                const rows = snapshot.docs.map((doc) => {
                    const data = doc.data();
                    console.log(data);
                    return [data.name, data.surname ,data.email, data.lvl, data.encargado, data.state];
                });
            
                rows.unshift(['Name', 'Surname', 'Email', 'Nivel',  'Encargado', 'Status']);
            
                snapshot.docChanges().forEach((change) => {
                    if (change.type === 'removed') {
                    }else{
                        sheets.spreadsheets.values.update({
                          spreadsheetId: sheetId,
                          range: doc.id+'!A1:F10000',
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

        }else{
            sheets.spreadsheets.get({
                spreadsheetId: sheetId,
            }, (err, res) => {
                if (err) return console.log('The API returned an error: ' + err);
                const id = res.data.sheets.find(s => s.properties.title === doc.id).properties.sheetId;
                console.log(id);
                //delete the sheet
                sheets.spreadsheets.batchUpdate({
                    spreadsheetId: sheetId,
                    resource: {
                        requests: [{
                            deleteSheet: {
                                sheetId: id
                            }
                        }]
                    }
                }, (err, res) => {
                    if (err) return console.log('The API returned an AAAAAA: ' + err);
                    console.log('Sheet deleted');
                });
            });
        }
    });
});
