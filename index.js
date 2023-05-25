// Javascript code to fill out certification form
// Edit execute command at bottom of file
// "node index.js" to run

// set the sender email address before running
const fromEmail = '';

// Set the environment variable $APP_PASSWORD before running
// This is required for authentication to gmail
const appPassword = process.env.APP_PASSWORD;

const readline = require('readline');
const { PDFDocument } = require('pdf-lib');
const { readFile, writeFile } = require('fs/promises');
const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

async function createPdf(username, cert_date, course, input='form/certificate_form.pdf') {
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
      return output;
  } catch (err) {
      console.log(err);
  }
};

function sendCertificate(filename, toEmail, fromE=fromEmail, passwd=appPassword) {
  const pdfFile = fs.readFileSync(filename);

  let transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: fromE,
        pass: passwd
    }
  });

  let mailOptions = {
    from: fromE,
    to: toEmail,
    subject: 'H2O.ai Certificate',
    text: 'Please find attached the certificate from H2O.ai.',
    attachments: [{
        filename: "certificate.pdf",
        content: pdfFile
    }]
  };

  return new Promise((resolve, reject) => {
    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        reject(error);
      } else {
        resolve(info.response);
      }
    });
  });
}

async function loopOverExcel(inputFile, theSheet, fromEmail, appPassword) {
  const workbook = xlsx.readFile(inputFile);
  const worksheet = workbook.Sheets[theSheet];

  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);

  for (let row of rows) {
    const [username, toEmail, course, date] = row;

    // check if any variable is undefined
    if (typeof username === 'undefined' || typeof toEmail === 'undefined' || typeof course === 'undefined' || typeof date === 'undefined') {
      console.log(`Skipping row ${row} because at least one field is undefined`);
      continue;
    }

    // check if any variable is empty or contains only whitespace
    if (!username.trim() || !toEmail.trim() || !course.trim() || !date.trim()) {
      console.log(`Skipping row ${row} because at least one field is empty`);
      continue;
    }

    try {
      const output = await createPdf(username, date, course);
      await sendCertificate(output, toEmail);
      console.log(`Certificate sent to ${toEmail} for ${username}`);
    } catch (error) {
      console.log(`Error generating or sending certificate to ${username}`, error);
    }
  }

  console.log('Certificates generated for all rows!');
}

// Try with test case
// createPdf("Amol Shanbhag", "20 April 2023", "Beginner's H2O Driverless AI Training")

// Try with entire spreadsheet 
// loopOverExcel("H2OWorld.xlsx", 'Sheet1', fromEmail, appPassword);


