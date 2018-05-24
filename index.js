const express = require("express");
const Excel = require("exceljs");
const request = require("request");

const http = require("http");
const fs = require("fs");
const path = require("path");

const app = express();
const port = process.env.PORT || 3000;

const FILEPATH = path.resolve(process.cwd(), 'sheet.xlsx');

// Use middleware to parse incoming requests containing JSON payloads and FORM data
app.use(express.json());

app.get('/', async function(req, res) {
    res.sendFile(FILEPATH);
});

async function getWorkbook(filePath) {
    try {
        let workbook = new Excel.Workbook();
        return await workbook.xlsx.readFile(filePath);
    } catch (error) {
        console.warn('Unable to get workbook:');
        console.error(error);
    }
}

async function downloadFile(downloadUrl, savePath) {
    // return a promise and resolve when download finishes
    return new Promise((resolve, reject) => {
        request
            .get(downloadUrl)
            .on('error', error => {
                reject(error);
            })

            .pipe(fs.createWriteStream(savePath))
            .on('finish', function() {
                resolve();
            })
            .on('error', error => {
                reject(error);
            });
    });
}

async function run() {
    try {
        const downloadUrl = 'http://www.calendarpedia.co.uk/download/calendar-2017-landscape-in-colour.xlsx';

        await downloadFile(downloadUrl, FILEPATH);
        const workbook = await getWorkbook(FILEPATH);

    } catch (error) {
        console.warn('Unable to run');
        console.error(error)
    }
}

run();

// Start web server
http.createServer(app).listen(port, () => {
    console.log(`Starting HTTP web server: http://localhost:${port}`);
});