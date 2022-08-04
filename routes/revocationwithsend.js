var express = require('express');
var router = express.Router();
const fs = require('fs');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const nodemailer = require("nodemailer");
const cors = require('cors');
require('log-timestamp');

var data = require('./data.js');
const emaildata = {};
const LETTERS_DIR = data.filePath;

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
    res.json({ message: 'revocation letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    var date1 = new Date();
    const DATE = dateFormat(date1, "dd-mmm-yyyy");


    var docDefinition = {
        pageSize: 'A4',
        pageOrientation: 'portrait',
        pageMargins: [50, 60, 50, 60],
        footer: {
            columns: [
                { text: data.footeroneline }
            ],
            style: 'superMargin'
        },

        content: [
            {
                alignment: 'justify',
                columns: [
                    {
                        image: 'coop.jpg',
                        width: 300
                    },
                    {
                        type: 'none',
                        alignment: 'right',
                        fontSize: 9,
                        ol: [
                            'The Co-operative Bank of Kenya Limited',
                            'Head Office',
                            'The Co-operative Bank of Kenya Limited',
                            'Co-operative Bank House',
                            'Haile Selassie Avenue',
                            'P.O.Box 48231-00100 GPO, Nairobi',
                            'Tel: (020) 3276100',
                            'Fax: (020) 2227747/2219831',
                            { text: 'Website: www.co-opbank.co.ke', color: 'blue', link: 'http://www.co-opbank.co.ke' }
                        ]
                    },
                ],
                columnGap: 10
            },
            '\nOur Ref: ' + letter_data.custnumber,
            '\n' + DATE,
            '\n' + letter_data.insuranceco,
            '' + letter_data.insuranceaddress,
            // '' + letter_data.postcode,
            ' Head Office',
            '\nNAIROBI',

            '\nAttn: Underwriting Manager',

            '\nDear Sir/Madam',
            {
                text: '\nRE: REVOCATION \n' + letter_data.custname + '\nPOLICY NO: ' + letter_data.policynumber,
                style: 'subheader'
            },
            {
                text: [
                    '\nWe herein notify you that the above applicant has regularized her Insurance Premium Finance (IPF) account, and advise you to reinstate the customer’s insurance policy.'
                ]
            },

            { text: '\nYour co-operation is highly appreciated. ', fontSize: 11, alignment: 'left' },

            { text: '\nYours Faithfully,' },
            { text: '\n\nDAVID MITHIA,                                                                           JAMES KARANJA', style: 'tableHeader' },
            { text: 'REMEDIAL CREDIT DEPARTMENT                                          FOR HEAD – MSME REMEDIAL CREDIT DEPARTMENT ', style: 'tableHeader' },
            { text: '\n\n\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },


            { text: '\nCc ' },
            { text: '' + letter_data.custname },
            { text: '' + letter_data.branchname }
        ],

        styles: {
            header: {
                fontSize: 18,
                bold: true,
                alignment: 'right',
                margin: [0, 190, 0, 80]
            },
            subheader: {
                fontSize: 12,
                bold: true,
                decoration: 'underline'
            },
            superMargin: {
                margin: [20, 0, 40, 0],
                fontSize: 8, alignment: 'center', opacity: 0.5
            },
            quote: {
                italics: true
            },
            small: {
                fontSize: 8
            },
            tableHeader: {
                bold: true,
                fontSize: 10,
                color: 'black'
            }
        },
        defaultStyle: {
            fontSize: 10
        }
    };



    var options = {
        // ..
    }

    var pdfDoc = printer.createPdfKitDocument(docDefinition, options);
    writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.accnumber + DATE + "revocation.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.accnumber + DATE + "revocation.pdf",
            filename: letter_data.accnumber + DATE + "revocation.pdf"
        });

        // send email
        emaildata.custname = letter_data.custname,
            emaildata.email = letter_data.insuranceemail,
            emaildata.policynumber = letter_data.policynumber,
            emaildata.path = LETTERS_DIR + letter_data.accnumber + DATE + "revocation.pdf",
            emaildata.cc = [letter_data.emailaddress];


        var transporter = nodemailer.createTransport({
            host: data.smtpserver,
            port: data.smtpport,
            secure: false, // upgrade later with STARTTLS
            tls: { rejectUnauthorized: false },
            debug: true,
            auth: {
                user: data.smtpuser,
                pass: data.pass
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
            from: data.smtpuser,
            to: emaildata.email,
            cc: emaildata.cc,
            subject: "IPF Reinstatement - " + emaildata.custname + " ( Policy No: " + emaildata.policynumber + " )",
            // text: "Text. ......",
            html: '<h5>Dear Sir/Madam:</h5>' +
                'Please find attached IPF reinstatement letter for the above customer.<br>' +
                '<p>Kindly note this is an automated delivery system; do not reply to this email address</p>' +
                '<br>' +
                'For any queries, kindly contact Customer Service on phone numbers: 0703027000/ 020 2776000 | SMS:16111 | <br>' +
                'Email: customerservice@co-opbank.co.ke | Twitter handle: @Coopbankenya | Facebook: Co-opBank Kenya | WhatsApp:0736690101<br>' +
                '<br>' +
                'Best Regards,<br>' +
                'Co-operative Bank of Kenya' +
                '<br> <br>',
            attachments: [
                {
                    path: emaildata.path
                }
            ]
        };

        // send email
        transporter.sendMail(mailOptions, (error, info) => {
            if (error) {
                console.log(error);
                res.json({
                    result: 'fail',
                    message: "email not sent"
                })
            }
            console.log("info > ", info)
            res.json({
                result: 'success',
                message: "email sent",
                info: info
            })
        })
    });

});

module.exports = router;
