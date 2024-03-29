var express = require('express');
var router = express.Router();
const nodemailer = require("nodemailer");
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const cors = require('cors');
var Minio = require("minio");
var minioClient = new Minio.Client({
    endPoint: process.env.MINIO_ENDPOINT || '127.0.0.1',
    port: process.env.MINIO_PORT ? parseInt(process.env.MINIO_PORT, 10) : 9005,
    useSSL: false,
    accessKey: process.env.ACCESSKEY || 'AKIAIOSFODNN7EXAMPLE',
    secretKey: process.env.SECRETKEY || 'wJalrXUtnFEMIK7MDENGbPxRfiCYEXAMPLEKEY'
});

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
    res.json({ message: 'Ipf cancellation letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    var date1 = new Date();
    const DATE = dateFormat(date1, "dd-mmm-yyyy");
    const rawaccnumber = letter_data.accnumber;
    const repodate = letter_data.daterepoissued;
    const repodate2 = dateFormat(repodate, "dd-mmm-yyyy");
    const first4 = rawaccnumber.substring(0, 9);
    const accnumber_masked = first4 + 'xxxxx';

    var DATEEXPIRY = dateFormat(date1.setDate(date1.getDate() + 30), "dd-mmm-yyyy");
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
            {
                style: 'tableExample',
                table: {
                    headerRows: 1,
                    widths: [250, '*'],
                    body: [
                        [{text: '', style: 'tableHeader'}, {text: '', style: 'tableHeader'}],
                        ['Date: ' + repodate2, 'Valid up to: ' + DATEEXPIRY],
                        ['To ', ': ' + letter_data.auctioneername],
                        ['Asset Finance Agreement No.', ': ' + letter_data.assetfaggnum],
                        ['Hirer’s Name ', ':     ' + letter_data.custname],
                        ['Unit Financed ', ':     ' + letter_data.vehiclemake + ' & ' + letter_data.vehiclemodel],
                        ['Registration No ', ':     ' + letter_data.vehicleregno],
                    ]
                },
                layout: 'noBorders',

            },
            '\n',

            {
                text: [
                    '\nAccording to our records, the monthly rental of the above Asset finance Agreement is now in arrears. The total amount due is ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.totalamount)).format('0,0.00'), fontSize: 10, bold: true },
                    '. ',
                ]

            },
            {
                text: [
                    '\nPlease approach the above named Hirer on our behalf and collect the total sum of ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.totalamount)).format('0,0.00'), fontSize: 10, bold: true },
                    ' plus your own charges, failing which you may take this letter as your authority to effect immediate re-possession',
                    ' of the above/equipment without further reference to us.',
                    { text: 'HIRER MUST MAKE PAYMENT VIDE CASH OR BY BANKER’S CHEQUE AS PERSONAL CHEQUE(S) WILL NOT BE ACCEPTED.', fontSize: 10, bold: true},
                    'From our records, we are able to give the following additional information regarding this Agreement, which may assist you in your task of locating the hirer and/or the motor vehicle/equipment:-',
                ]
            },
            '\n',

            {
                table: {
                    headerRows: 1,
                    widths: [200, '*'],
                    body: [
                        [{text: '', style: 'tableHeader'}, {text: '', style: 'tableHeader'}],
                        ['Postal Address', ':      ' + letter_data.postaladdress || 'N/A'],
                        ['Telephone', ':     ' + letter_data.celnumber || 'N/A'],
                        ['Physical Address/Location', ':     ' + letter_data.place || 'N/A'],
                        ['Type of Business', ':      ' + letter_data.typeofbusiness || 'N/A'],
                        ['Bankers and Branch', ':     ' + letter_data.branchname || 'N/A'],
                        ['Purpose of Vehicle', ':     ' + letter_data.purposeofvehicle || 'N/A'],
                        ['Guarantors', ':     ' + letter_data.guarantors || 'N/A'],
                        ['Guarantors Address', ':     ' + letter_data.guarantorsaddress || 'N/A'],
                        ['Chassis No.', ':     ' + letter_data.chassisnumber || 'N/A'],
                        ['Engine No.', ':     ' + letter_data.engineno || 'N/A'],
                        ['Any other information', '     ' + letter_data.anyotherinfo || 'N/A']
                    ]
                },
                layout: 'noBorders',

            },

            { text: '\n\nVehicle tracked by: ' + letter_data.trackingcompany, fontSize: 10, alignment: 'left' },

            { text: '\nYours Faithfully,' },



            { text: '\n\nAUTHORISED SIGNATORY,                                                                           AUTHORISED SIGNATORY', style: 'tableHeader' },
            { text: '\n\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },



            { text: '\nCc ' },
            { text: '' + letter_data.custname },
            { text: '' + letter_data.postaladdress }
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
    writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "repossession.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        // res.json({
        //     result: 'success',
        //     message: LETTERS_DIR + letter_data.accnumber + DATE + "repossession.pdf",
        //     filename: letter_data.accnumber + DATE + "repossession.pdf"
        // })
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "repossession.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "repossession.pdf"
        var metaData = {
            'Content-Type': 'text/html',
            'Content-Language': 123,
            'X-Amz-Meta-Testing': 1234,
            'example': 5678
        }
        minioClient.fPutObject(bucket, savedfilename, filelocation, metaData, function (error, objInfo) {
            if (error) {
                console.log(error);
                res.status(500).json({
                    success: false,
                    error: error.message
                })
                // deleteFile(filelocation);
            }
            res.json({
                result: 'success',
                message: LETTERS_DIR + accnumber_masked + DATE + "repossession.pdf",
                filename: accnumber_masked + DATE + "repossession.pdf",
                savedfilename: savedfilename,
                objInfo: objInfo
            })
            // deleteFile(filelocation);
        });
        //save to mino end

        // send email
        emaildata.customerName = letter_data.custname,
            emaildata.email = letter_data.auctioneeremail,
            emaildata.vehicleRegnumber = letter_data.vehicleregno,
            emaildata.path = LETTERS_DIR + accnumber_masked + DATE + "repossession.pdf";
           // emaildata.cc = [letter_data.emailaddress];


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
            subject: "Repossession Order - " + emaildata.customerName + " VEHICLE REG NO.: " + emaildata.vehicleRegnumber,
            // text: "Text. ......",
            html: '<h5>Dear Sir/Madam:</h5>' +
                'Please find attached repossession order for the above customer.<br>' +
                '<p>Kindly note this is an automated delivery system; do not reply to this email address</p>' +
                '<br>' +
                //'For any queries, kindly contact Customer Service on phone numbers: 0703027000/ 020 2776000 | SMS:16111 | <br>' +
                //'Email: customerservice@co-opbank.co.ke | Twitter handle: @Coopbankenya | Facebook: Co-opBank Kenya | WhatsApp:0736690101<br>' +
                '<br>' +
                'Best Regards,<br>' +
                letter_data.rmname + ',<br>' +
                'Co-operative Bank of Kenya,<br>' +
                'Tel: ' + letter_data.rmtelnumber + ',<br>' +
                'Email: ' + letter_data.rmemail +
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
});
// function deleteFile(req) {
//     fs.unlink(req, (err) => {
//         if (err) {
//             console.error(err)
//             return
//         }
//         //file removed
//     })
// }

module.exports = router;
