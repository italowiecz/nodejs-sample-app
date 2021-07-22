const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const moment = require('moment');

async function xlsxExtractData() {
    const workbook = new ExcelJS.Workbook();
    return new Promise(resolve => {
        workbook.xlsx.readFile('Sample.xlsx').then(function () {
            var worksheet = workbook.getWorksheet('Sheet1');
            var obj = {};
            worksheet.eachRow({ includeEmpty: false }, function (row) {
                var key = moment(row.values[2]).format('DD/MM/YYYY');
                obj[key] = [];
                var rowData = {
                    name: row.values[1],
                    email: row.values[3].text
                };
                obj[key].push(rowData);
            }
            );
            console.log(obj);
            resolve(obj);
        });
    });
}

function checkWhoIsOnCall(xlsxObject) {
    return new Promise((resolve, reject) => {
        var currentDate = moment(new Date()).format('D/MM/YYYY');
        var onCallPerson = xlsxObject[currentDate];
        if (onCallPerson != null) {
            console.log('On call entity:', Object.values(onCallPerson)[0]);
            resolve(Object.values(onCallPerson)[0]);
        } else {
            reject('There is nobody on call today!');
        }
    });
}

function sendOnCallReminder(onCallPerson){
    var transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'email@gmail.com',
            pass: 'PASSWORD'
        }
    });
    
    var mailOptions = {
        from: 'email@gmail.com',
        to: onCallPerson.email,
        subject: 'REMINDER: You are on call today!',
        text: 'This is a friendly reminder of your on call agenda.'
    };

    if(onCallPerson.email != 'null'){
        transporter.sendMail(mailOptions, function (error, info) {
            if (error) {
                console.log(error);
            } else {
                console.log('Email sent: ' + info.response);
            }
        });
    }
}

xlsxExtractData().then(obj => {checkWhoIsOnCall(obj)
                 .then(onCallPerson => {sendOnCallReminder(onCallPerson)})
                 .catch(err => {console.log(err)});});

    

  


    



