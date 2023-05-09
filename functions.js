const { PDFDocument } = require('pdf-lib');
const { readFile, writeFile } = require('fs/promises');
const fs = require('fs');
const csv = require('csv-parser');
const xlsx = require('xlsx');
//const createPdf = require('./createPdf');

async function createPdf(username, cert_date, course, input='certificate_form.pdf') {
    try {
        const pdfDoc = await PDFDocument.load(await readFile(input));
        const form = pdfDoc.getForm();
        form.getTextField('text_username').setText(username);
        form.getTextField('text_coursework').setText(course);
        form.getTextField('date').setText(cert_date);
        form.flatten();
        const pdfBytes = await pdfDoc.save();

        // create unique output filename
        const d = new Date();
        var str = '';
        str += username.replace(/ /g, '');
        str += '_';
        str += d.getTime();
        var output = str += '.pdf';
        await writeFile(output, pdfBytes);
        console.log(output);
        console.log('PDF Created');
    } catch (err) {
        console.log(err);
    }
}

async function loopOverCsv(inputFile) {
    fs.createReadStream(inputFile)
      .pipe(csv())
      .on('data', (row) => {
        try {
          const username = row.username;
          const date = row.date;
          const course = row.course;

          createPdf(username, date, course);
        } catch (err) {
          console.error(`Error generating certificate for row ${row}`, err);
        }
      })
      .on('end', () => {
        console.log('Certificates generated for all rows!');
      })
}

async function loopOverXlsx(inputFile) {
    const workbook = xlsx.readFile(inputFile);
    const worksheet = workbook.Sheets['Sheet1'];

    const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);

    rows.forEach((row) => {
      try {
        const [username, date, course] = row;

        createPdf(username, date, course);
      } catch (err) {
        console.error(`Error generating certificate for row ${row}`, err);
      }
    });

    console.log('Certificates generated for all rows!');
}

async function loopOverExcel(inputFile, date, course) {
    const workbook = xlsx.readFile(inputFile);
    const worksheet = workbook.Sheets['Sheet1'];

    const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);

    rows.forEach((row) => {
      try {
        const [first, last] = row;
	var username = first + ' ' + last;
        createPdf(username, date, course);
      } catch (err) {
        console.error(`Error generating certificate for row ${row}`, err);
      }
    });

    console.log('Certificates generated for all rows!');
}

export { createPdf, loopOverCsv, loopOverXlsx, loopOverExcel }


// app password
// vfmenvhsnqplrvzl