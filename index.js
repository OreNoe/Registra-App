const {google} = require('googleapis');
const admin = require('firebase-admin');
const express = require('express');
var request = require('request');
const serviceAccount = require('./serviceAccountKey.json');
const https = require('https');

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

const app = express();
app.use(express.json());
const port = process.env.PORT || 8080;

app.listen(port, () => {
    console.log(`Listening on port ${port}`);
});

//create post request to add new event to google sheets so create a new sheet, the parameter is the event name. Also set the header row in the new sheet. 
//Send a 200 response if successful, 500 if not.
app.post('/addEvent', (req, res) => {
    if (!req.body.eventName) return res.sendStatus(400);
    const eventSheetName = req.body.eventName.replace(/\s+/g, ''); // Remove whitespace from the event name
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
    //create the sheet
    sheets.spreadsheets.batchUpdate(eventSheet, (err, res) => {
        if (err) return console.log('Error agregando evento: ' + err);
        //console.log(res);
        //set the header row
        sheets.spreadsheets.values.update({
            spreadsheetId: sheetId,
            range: eventSheetName + '!A:F',
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [['Nombre', 'Apellido', 'Email', 'Nivel', 'Status', 'Asesor']],
            },
        }, (err, res) => {
            if (err) return console.log('Error agregando encabezado: ' + err);
        });

        sheets.spreadsheets.get({
            spreadsheetId: sheetId,
        }, (err, res) => {
            if (err) return console.log('Error obteniendo evento a agregar: ' + err);
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
                        }
                    }],
                },
            };
            sheets.spreadsheets.batchUpdate(eventS, (err, res) => {
                if (err) return console.log('Error formateando evento agregado: ' + err);
                //console.log(res);
            });
        });
    });
    res.sendStatus(200);
});
app.post('/updateUser', (req, res) => {
    const eventSheetName = req.body.eventName.replace(/\s+/g, ''); // Remove whitespace from the event name
    const body = req.body;
    switch (body.lvl) {
        case '1':
            body.lvl = 'General';
            break;
        case '2':
            body.lvl = 'VIP';
            break;
        case '3':
            body.lvl = 'Invitado';
            break;
    }
    //get the sheet id
    sheets.spreadsheets.get({
        spreadsheetId: sheetId,
    }, (err, res) => {
        if (err) return console.log('Error obteniendo evento a agregar usuario: ' + err);
        //add user to sheet
        sheets.spreadsheets.values.update({
            spreadsheetId: sheetId,
            range: eventSheetName + '!A:F',
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[body.name, body.surname, body.email, body.lvl, body.status, body.encargado]],
            },
        }, (err, res) => {
            if (err) return console.log('Error agregando usuario: ' + err);
        });
        sheets.spreadsheets.get({
                spreadsheetId: sheetId,
            }, (err, res) => {
                if (err) return console.log('Error obteniendo evento a actualizar: ' + err);
                const id = res.data.sheets.find(s => s.properties.title === eventSheetName).properties.sheetId;
                //set the conditional format

                //setBasicFilter to the sheet that orders by the status column
                sheets.spreadsheets.batchUpdate({
                    spreadsheetId: sheetId,
                    resource: {
                        requests: [{
                            setBasicFilter: {
                                filter: {
                                    range: {
                                        sheetId: id,
                                        startRowIndex: 0,
                                        endRowIndex: 1000,
                                        startColumnIndex: 0,
                                        endColumnIndex: 6,
                                    },
                                    sortSpecs: [{
                                        dimensionIndex: 4,
                                        sortOrder: 'ASCENDING',
                                    }],
                                },
                            },
                        }],
                    },
                }, (err, res) => {
                    if (err) return console.log('Error ordenando usuario: ' + err);
                    //console.log(res);
                });
            });
    });
    res.sendStatus(200);
});

//create post request to add new user to google sheets so add a new row to the sheet with the event name. The parameters are the event and a user object with the user data.
//Send a 200 response if successful, 500 if not.
app.post('/addUserToEvent', (req, res) => {
    const eventSheetName = req.body.eventName.replace(/\s+/g, ''); // Remove whitespace from the event name
    const body = req.body;
    //get the sheet id
    sheets.spreadsheets.get({
        spreadsheetId: sheetId,
    }, (err, res) => {
        if (err) return console.log('Error obteniendo evento a agregar usuario: ' + err);
        //add user to sheet
        switch (body.lvl) {
            case '1':
                body.lvl = 'General';
                break;
            case '2':
                body.lvl = 'VIP';
                break;
            case '3':
                body.lvl = 'Invitado';
                break;
        }

        sheets.spreadsheets.values.append({
            spreadsheetId: sheetId,
            range: eventSheetName + '!A:F',
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[body.name, body.surname, body.email, body.lvl, body.status, body.encargado]],
            },
        }, (err, res) => {
            if (err) return console.log('Error agregando usuario: ' + err);
            //console.log(res);
        });
    });

    res.sendStatus(200);
});

//create post request to delete a sheet with the event name. The parameter is the event name.
//Send a 200 response if successful, 500 if not.
app.post('/deleteEvent', (req, res) => {
    const eventSheetName = req.body.eventName.replace(/\s+/g, ''); // Remove whitespace from the event name
    //get the sheet id
    sheets.spreadsheets.get({
        spreadsheetId: sheetId,
    }, (err, res) => {
        if (err) return console.log('Error obteniendo evento a eliminar: ' + err);
        const id = res.data.sheets.find(s => s.properties.title === eventSheetName).properties.sheetId;
        
        sheets.spreadsheets.batchUpdate({
            spreadsheetId: sheetId,
            resource: {
                requests: [{
                    deleteSheet: {
                        sheetId: id,
                    },
                }],
            },
        }, (err, res) => {
            if (err) return console.log('Error eliminando evento: ' + err);
        });
    });

    res.sendStatus(200);
});

app.post('/deleteUser', (req, res) => {
    const eventSheetName = req.body.eventName.replace(/\s+/g, ''); // Remove whitespace from the event name
    //get user row
    sheets.spreadsheets.values.get({
        spreadsheetId: sheetId,
        range: eventSheetName + '!A:F',
    }, (err, res) => {
        if (err) return console.log('Error obteniendo usuario: ' + err);
        const values = res.data.values;
        if (values.length) {
            const row = values.find(r => r[2] === body.email);
            const index = values.indexOf(row);
            if (index === -1) return console.log('Error: No se encontró el usuario');
            //delete user row
            sheets.spreadsheets.values.batchClear({
                spreadsheetId: sheetId,
                ranges: [eventSheetName + '!A' + (index + 1) + ':F' + (index + 1)],
            }, (err, res) => {
                if (err) return console.log('Error eliminando usuario: ' + err);
                //console.log(res);
            });
            //move all rows up
            sheets.spreadsheets.get({
                spreadsheetId: sheetId,
            }, (err, res) => {
                if (err) return console.log('Error obteniendo evento a eliminar: ' + err);
                const id = res.data.sheets.find(s => s.properties.title === eventSheetName).properties.sheetId;
                sheets.spreadsheets.values.batchUpdate({
                    spreadsheetId: sheetId,
                    resource: {
                        requests: [{
                            deleteRange: {
                                range: {
                                    sheetId: id,
                                    startRowIndex: index,
                                    endRowIndex: index + 1,
                                    startColumnIndex: 0,
                                    endColumnIndex: 6,
                                },
                                shiftDimension: 'ROWS',
                            },
                        }],
                    },
                }, (err, res) => {
                    if (err) return console.log('Error moviendo filas: ' + err);
                    //console.log(res);
                });
            });
        }
    });
    res.sendStatus(200);
});

app.get('/', (req, res) => {
    res.send('Server awake!' + new Date().toLocaleString());
});