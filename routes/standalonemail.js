var express = require('express');
var router = express.Router();
const nodemailer = require("nodemailer");
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const cors = require('cors');

var data = require('./data.js');

const LETTERS_DIR = data.filePath;
const emaildata = {};
var fonts = {
    Roboto: {
        normal: 'fonts/Roboto-Regular.ttf',
        bold: 'fonts/Roboto-Medium.ttf',
        italics: 'fonts/Roboto-Italic.ttf',
        bolditalics: 'fonts/Roboto-MediumItalic.ttf'
    }
};

var PdfPrinter = require('pdfmake');
var printer = new PdfPrinter(fonts);

router.use(bodyParser.urlencoded({
    extended: true
}));

router.use(bodyParser.json());
router.use(cors());


router.get('/', function (req, res) {
    res.json({ message: 'Email sending is Ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    console.log(letter_data)
    var date1 = new Date();
    const DATE = dateFormat(date1, "dd-mmm-yyyy");



        // send email
        emaildata.custname = letter_data.clientname,
            emaildata.email = letter_data.emailaddress,
            emaildata.branchemail = 'E-Collect <ecollect@co-opbank.co.ke>',
            emaildata.policynumber = letter_data.policynumber,
            // emaildata.path = LETTERS_DIR + letter_data.accnumber + DATE + "ipfcancellation.pdf",
            emaildata.cc = letter_data.cc,
        emaildata.emailcontent = letter_data.emailcontent,
            emaildata.subject = letter_data.subject


    let transporter = nodemailer.createTransport({
        host: 'smtp.gmail.com',
        port: 587,
        authentication:'plain',
        secure: false, // true for 465, false for other ports
        auth: {
            user: 'allanmaroko10',
            pass: 'Vipermarox411'
        }
    });


    // verify connection configuration
        transporter.verify(function (error, success) {
            if (error) {
                console.log(error);
            } else {
                console.log("Server is ready to take our messages");
            }
        });


        var mailOptions = {
            from: 'ecollect@co-opbank.co.ke',
            to: emaildata.email,
            cc: emaildata.cc,
            subject: emaildata.subject,
            // text: "Text. ......",
            html: emaildata.emailcontent,
            attachments: [
                {
                    path: emaildata.path
                }
            ]
        };
       // start to check if there is an attachment
        if (mailOptions.attachments.path === undefined) {
            console.log('no attachment detected');
            delete mailOptions['attachments'];
        }

        // end of checking if there is an attachment




        // send email
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                console.log(error);
                res.json({
                    result: 'fail',
                    message: "message not sent"
                })
            }
            console.log("info > ", info)
            res.json({
                result: 'success',
                message: "message sent",
                info: info
            })
        })
    });

module.exports = router;
