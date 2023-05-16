const appPassword = 'vfmenvhsnqplrvzl';
const readline = require('readline');

const { PDFDocument } = require('pdf-lib');
const { readFile, writeFile } = require('fs/promises');
const fs = require('fs');
const xlsx = require('xlsx');

const nodemailer = require('nodemailer');
const fromEmail = 'david.whiting@h2o.ai';

// set the environment variable $APP_PASSWORD
// const appPassword = process.env.APP_PASSWORD;

async function createCbaPdf(username, cert_date, input='form/CBA-H2O_form.pdf') {
  try {
      const pdfDoc = await PDFDocument.load(await readFile(input));
      const form = pdfDoc.getForm();
      form.getTextField('text_username').setText(username);
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

function sendCbaCert(filename, toEmail, fromE=fromEmail, passwd=appPassword) {
  const pdfFile = fs.readFileSync(filename);
  const pdfAttach = fs.readFileSync("form/cba_boot_camp.pdf");

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
    text: 'Please find attached the certificate from the H2O.ai bootcamp.',
    attachments: [
      {
        filename: "certificate.pdf",
        content: pdfFile
      },
      {
        filename: "bootcamp.pdf",
        content: pdfAttach
      }
    ]
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

async function loopOverCbaExcel(inputFile, theSheet, fromEmail, appPassword) {
  const workbook = xlsx.readFile(inputFile);
  const worksheet = workbook.Sheets[theSheet];

  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 }).slice(1);

  for (let row of rows) {
    const [username, toEmail, date] = row;

    // check if any variable is undefined
    if (typeof username === 'undefined' || typeof toEmail === 'undefined' || typeof date === 'undefined') {
      console.log(`Skipping row ${row} because at least one field is undefined`);
      continue;
    }

    // check if any variable is empty or contains only whitespace
    if (!username.trim() || !toEmail.trim() || !date.trim()) {
      console.log(`Skipping row ${row} because at least one field is empty`);
      continue;
    }

    try {
      const output = await createCbaPdf(username, date);
      await sendCbaCert(output, toEmail);
      console.log(`Certificate sent to ${toEmail} for ${username}`);
    } catch (error) {
      console.log(`Error generating or sending certificate for ${username}`, error);
    }
  }

  console.log('Certificates generated for all rows!');
}


// Try with test case

// createCbaPdf("Ben Simmonds", "20 April 2023");
loopOverCbaExcel("CBA.xlsx", 'Final', fromEmail, appPassword);

