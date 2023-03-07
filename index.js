const {google} = require('googleapis');
const admin = require('firebase-admin');
const fs = require('fs');
const ptp = require('node-printer');

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
            return [data.name];
        });

        rows.unshift(['Name']);

        if (change.type === 'added') {
            const event = doc.data();
            const eventSheetName = event.name;
            const eventSheet = {
                spreadsheetId: sheetId,
                resource: {
                    requests: [{
                        addSheet: {
                            properties: {
                                title: eventSheetName,
                            },
                        },
                    
                    }],
                },
            };

            sheets.spreadsheets.batchUpdate(eventSheet, (err, res) => {
                if (err) return console.log('The API returned an error: ' + err);
                //console.log(res);
                sheets.spreadsheets.get({
                    spreadsheetId: sheetId,
                }, (err, res) => {
                    if (err) return console.log('The API returned an error: ' + err);
                    const id = res.data.sheets.find(s => s.properties.title === eventSheetName).properties.sheetId;
    
                    const eventS = {
                        spreadsheetId: sheetId,
                        resource: {
                            requests: [{
                                addConditionalFormatRule: {
                                    rule: {
                                        ranges: [{
                                            sheetId: id,
                                            startRowIndex: 0,
                                            endRowIndex: 1000,
                                            startColumnIndex: 4,
                                            endColumnIndex: 6,
                                        }],
                                        booleanRule: {
                                            condition: {
                                                type: 'TEXT_EQ',
                                                values: [{
                                                    userEnteredValue: 'true',
                                                }],
                                            },
                                            format: {
                                                backgroundColor: {
                                                    red: 0.0,
                                                    green: 1.0,
                                                    blue: 0.0,
                                                },
                                            },
                                        },
                                    },
                                    index: 0,
                                },
                            }],
                        },
                    };
                    sheets.spreadsheets.batchUpdate(eventS, (err, res) => {
                        if (err) return console.log('The API returned an error: ' + err);
                        //console.log(res);
                    });
                });
            });

            //get the sheet id
            


        }else if (change.type === 'modified') {
            const event = doc.data();
            const eventSheetName = event.name;
            const eventSheet = {
                spreadsheetId: sheetId,
                resource: {
                    requests: [{
                        updateSheetProperties: {
                            properties: {
                                title: eventSheetName,
                            },
                            fields: 'title',
                        },
                    }],
                },
            };
            sheets.spreadsheets.batchUpdate(eventSheet, (err, res) => {
                if (err) return console.log('The API returned an error: ' + err);
                //console.log(res);
            });
        }else if (change.type === 'removed') {
            sheets.spreadsheets.get({
                spreadsheetId: sheetId,
            }, (err, res) => {
                if (err) return console.log('The API returned an error: ' + err);
                const id = res.data.sheets.find(s => s.properties.title === doc.id).properties.sheetId;
                //console.log(id);
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
                    if (err) return console.log('The API returned an error: ' + err);
                    //console.log(res);
                });
            });
        }
    });
});


db.collection('events').onSnapshot(snapshot => {
    snapshot.docChanges().forEach(change => {
        const doc = change.doc;
        db.collection('events').doc(doc.id).collection('users').onSnapshot(snapshot => {
            const rows = snapshot.docs.map(doc => {
                const data = doc.data();
                console.log(data);
                return [data.name, data.surname, data.email, data.lvl, data.encargado, data.status];
            });

            rows.unshift(['Name', 'Surname', 'Email', 'Nivel', 'Encargado', 'Status']);

            if (change.type === 'added') {
                sheets.spreadsheets.values.update({
                    spreadsheetId: sheetId,
                    range: `${doc.id}!A1`,
                    valueInputOption: 'USER_ENTERED',
                    resource: {
                        values: rows,
                    },
                }, (err, res) => {
                    if (err) return console.log('The API returned an error: ' + err);
                    //console.log(res);
                });
            }else if (change.type === 'modified') {
                sheets.spreadsheets.values.update({
                    spreadsheetId: sheetId,
                    range: `${doc.id}!A1`,
                    valueInputOption: 'USER_ENTERED',
                    resource: {
                        values: rows,
                    },
                }, (err, res) => {
                    if (err) return console.log('The API returned an error: ' + err);
                    //console.log(res);
                });
                //send to printer the name of the user
                const user = doc.data();
                const name = user.name;
                //printName(name);

                
            }else if (change.type === 'removed') {
                sheets.spreadsheets.values.clear({
                    spreadsheetId: sheetId,
                    range: `${doc.id}!A1`,
                }, (err, res) => {
                    if (err) return console.log('The API returned an error: ' + err);
                    //dconsole.log(res);
                });
            }
        });
    });
});
/*
function printName(name) {
    const doc = new PDFDocument();
    doc.pipe(fs.createWriteStream('name.pdf'));
    doc.fontSize(60);
    doc.text(name, 0, 0);
    doc.end();

    printer.getDefaultPrinter().then(printerName => {
        printer
            .print('name.pdf', {
                printer: printerName.name,
                paperSize: 'A9',
                
            })
            .then(jobID => console.log('Printed to printer with ID: ', jobID))
            .catch(err => console.log('Error: ', err));
    });
}*/
