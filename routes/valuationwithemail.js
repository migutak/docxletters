var express = require('express');
var router = express.Router();
const nodemailer = require("nodemailer");
const fs = require('fs');
require('log-timestamp');
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
    res.json({ message: 'Ipf cancellation letter is ready!' });
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
            '\n\nOur Ref: MAO/'+letter_data.rrocode+'/CMD/'+letter_data.custnumber+'/2020',
            '',
            '\n' + DATE,
            '\n' + letter_data.serviceprovider,
            '' + letter_data.serviceprovideraddress,
            '' + letter_data.postcode,

            '\nDear Sir/Madam',
            {
                text: '\nRE: VALUATION ON MOTOR VEHICLE ',
                style: 'subheader'
            },
            {
                text: [
                    '\nPlease let us have valuation report on the below motor vehicle which we hold as security against a loan facility advanced to ',
                    { text: letter_data.custname, fontSize: 10, bold: true },
                    ' by our ',
                    { text: 'Retail Asset Finance Department.', fontSize: 10, bold: true },
                    ' Ensure that you also test the vehicle to see if it has any mechanical problems. '
                ]
            },
            '\nThereafter proceed to forward your fee note to the attention of the undersigned for settlement. ',
            {
                text: [
                    '\nPlease ensure to submit a complete report ',
                    { text: ' within 7 days ', fontSize: 10, bold: true },
                    ' from the date of our instructions. '
                ]
            },

            '\nYour prompt action will be highly appreciated.',
            '\n',
            {
                alignment: 'justify',
                fontSize: 10,
                table: {
                    body: [
                        [{ text: '', style: 'tableHeader', alignment: 'left' }, { text: 'VEHICLE DESCRIPTION', style: 'tableHeader', alignment: 'left' }, { text: 'REG. NO.', style: 'tableHeader', alignment: 'left' }, { text: 'REGISTERED OWNER', style: 'tableHeader', alignment: 'left' }, { text: 'STORAGE', style: 'tableHeader', alignment: 'left' }, { text: 'CONTACT PERSON', style: 'tableHeader', alignment: 'left' }],
                        ['1', letter_data.vehicledesc, letter_data.regnumber, letter_data.regowner, letter_data.yard, letter_data.serviceprovidercontact]
                    ]
                },
                layout: {
                    fillColor: function (rowIndex, node, columnIndex) {
                        return (rowIndex % 2 === 0) ? '#CCCCCC' : null;
                    }
                }
            },


            { text: '\n\n\nYours Faithfully,' },
            { text: '\n\nMitchelle Odhiambo,                                                                                      Charles Otieno,', style: 'tableHeader' },
            { text: 'Remedial Valuations Officer.                                                                        Manager-Remedial Service Providers.', style: 'tableHeader' },
            { text: 'Credit Management Div.                                                                                Credit Management Div.', style: 'tableHeader' },
            { text: '\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },

            { text: '\nCc: ' + letter_data.yard},
            { text: [
                '\nThe above valuers has been instructed by the bank to value Motor vehicle  ',
                { text: letter_data.regnumber + ' ' + letter_data.vehicledesc, fontSize: 10, bold: true },
                ' in the customer’s name and of Co-operative Bank of Kenya. Kindly accord them any assistance that they may require.'
            ] },

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
    writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.accnumber + DATE + "valuation.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.accnumber + DATE + "valuation.pdf",
            filename: letter_data.accnumber + DATE + "valuation.pdf"
        })

        // send email
            emaildata.path = LETTERS_DIR + letter_data.accnumber + DATE + "valuation.pdf"


        /*let transporter = nodemailer.createTransport({
            host: 'smtp.ethereal.email',
            port: 587,
            secure: false, // true for 465, false for other ports
            auth: {
                user: 'abdiel.funk26@ethereal.email',
                pass: '3HWTWq9gm2TgDtssNP'
            }
        });*/

        let transporter = nodemailer.createTransport({
            host: '192.168.0.39',
            port: 25,
            secure: false, // true for 465, false for other ports
            tls: { rejectUnauthorized: false },
            debug: true
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
            from: 'Collection Support <collectionssupport@co-opbank.co.ke>',
            to: letter_data.valueremail,
            subject: "Instruction to investigate - " + letter_data.custname,
            // text: "Text. ......",
            html: '<h5>Dear Sir/Madam:</h5>' +
                'Please find attached valuation instruction  for the above customer.<br>' +
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

module.exports = router;
