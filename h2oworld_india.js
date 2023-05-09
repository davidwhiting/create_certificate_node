const { PDFDocument } = require('pdf-lib');
const { readFile, writeFile } = require('fs/promises');
const fs = require('fs');
const xlsx = require('xlsx');

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

      // return the output filename
      // return output;
  } catch (err) {
      console.log(err);
  }
}

// import { createPdf, loopOverExcel } from "./functions.js";




var date = '19 April 2023';
var course = 'H2O World India: Hands-on Driverless AI Training';


async function loopOverExcel_old(inputFile, date, course) {
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

async function loopAndMailOverExcel(inputFile, theSheet) {
  const workbook = xlsx.readFile(inputFile);
  const worksheet = workbook.Sheets[theSheet];

  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);

  rows.forEach((row) => {
    try {
      const [username, email, course, date] = row;
      createPdf(username, date, course);
    } catch (err) {
      console.error(`Error generating certificate for row ${row}`, err);
    }
  });

  console.log('Certificates generated for all rows!');
}

async function loopAndMailOverExcel(inputFile, theSheet) {
  const workbook = xlsx.readFile(inputFile);
  const worksheet = workbook.Sheets[theSheet];

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



// loopOverExcel('Registrants.xlsx', date, course);

// create reusable transporter object using the default SMTP transport

// var username = "David Whiting";
// var certFile = '';
// certFile += username.replace(/ /g, '');
// certFile += '.pdf';

myCertFilename = createPdf("Abhinav Raja Raizada", "20 April 2023", "Beginner's H2O Driverless AI Training");
console.log(myCertFilename);