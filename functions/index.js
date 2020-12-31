const functions = require('firebase-functions');
const os = require('os');
const fs = require('fs');
const path = require('path');
const Busboy = require('busboy');
const admin = require('firebase-admin');
const XLSX = require('xlsx');


admin.initializeApp();


exports.readExcel = functions.https.onRequest(async (req, res) => {
    try {
        const busboy = new Busboy({ headers: req.headers });
        const tmpdir = os.tmpdir();

        const fields = {};

        const uploads = {};

        busboy.on('field', (fieldname, val) => {
            console.log(`Processed field ${fieldname}: ${val}.`);
            fields[fieldname] = val;
        });

        const fileWrites = [];
        const fileNames = {};

        busboy.on('file', (fieldname, file, filename) => {
            console.log(`Processed file ${filename}`);
            fileNames[fieldname] = filename;
            const filepath = path.join(tmpdir, filename);
            uploads[fieldname] = filepath;

            const writeStream = fs.createWriteStream(filepath);
            file.pipe(writeStream);
            const promise = new Promise((resolve, reject) => {
                file.on('end', () => {
                    writeStream.end();
                });
                writeStream.on('finish', resolve);
                writeStream.on('error', reject);
            });
            fileWrites.push(promise);
        });
        busboy.on('finish', async () => {
            await Promise.all(fileWrites);
            // upload file to firebase storage

            var workbook = XLSX.readFile(uploads['file']);
            var sheet_name_list = workbook.SheetNames;
            for (const y of sheet_name_list) {
                var worksheet = workbook.Sheets[y];
                var headers = {};
                var data = [];
                for (z in worksheet) {
                    if (z[0] === '!') continue;
                    //parse out the column, row, and value
                    var tt = 0;
                    for (var i = 0; i < z.length; i++) {
                        if (!isNaN(z[i])) {
                            tt = i;
                            break;
                        }
                    }
                    var col = z.substring(0, tt);
                    var row = parseInt(z.substring(tt));
                    var value = worksheet[z].v;

                    //store header names
                    if (row === 1 && value) {
                        headers[col] = value;
                        continue;
                    }

                    if (!data[row]) data[row] = {};
                    data[row][headers[col]] = value;
                }
                //drop those first two rows which are empty
                data.shift();
                data.shift();
                for (const el of data) {
                    const doc = admin.firestore().collection('SpreadsheetData').doc();
                    // eslint-disable-next-line no-await-in-loop
                    await doc.set(el)
                }
                // console.log(data);
            }

            for (const file in uploads) {
                fs.unlinkSync(uploads[file]);
            }

            res.status(200).send();
        });

        busboy.end(req.rawBody);
    } catch (e) {
        res.status(422).send(
            {
                "title": "Something went wrong",
                "message": e.toString()
            }
        )
    }

});