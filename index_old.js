//const fs = require('fs');
//const csv = require('csv-parser');
const createPdf = require('./createPdf');

var date = '19 April 2023';
var course = 'H2O World India: Hands-on Driverless AI';
var username = 'David Whiting';

createPdf(username, date, course);
