const { PDFDocument } = require('pdf-lib');
const { readFile, writeFile } = require('fs/promises');

var HAMC = 'H2O.ai University: Managed AI Cloud 101';
var HAIC = 'H2O.ai University: Hybrid AI Cloud 101';

async function createPdf(username, cert_date, course=HAMC, input='certificate_form.pdf') {
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
        str += d.getTime();
        str += '_';
        str += username.replace(/ /g, '');
        var output = str += '.pdf';
        await writeFile(output, pdfBytes);
        console.log(output);
        console.log('PDF Created');
    } catch (err) {
        console.log(err);
    }
}

var username = 'Eli Schultz';
var completion = '10/06/2022';

// Examples
createPdf('John Doe', '09/13/2022');
createPdf('Jane Doe', '10/06/2022', 'Super Partner Certification 101');

