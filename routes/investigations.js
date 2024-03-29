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
const EMAIL = data.SENDEMAILURL; // 'http://172.16.204.71:8005/demandemail/email';
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
    res.json({ message: 'investigators letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    var date1 = new Date();
    const DATE = dateFormat(date1, "dd-mmm-yyyy");

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
            '\n\nOur Ref: CMD/PL/2020',
            '',
            '\n' + DATE,
            '\n' + letter_data.serviceprovider,
            '' + letter_data.serviceprovideraddress,
            '' + letter_data.postcode,
      
            '\nDear Sir/Madam',
            {
                text: '\nRE: INVESTIGATION OF UNTRACEABLE AND MIGRATED ACCOUNTS',
                style: 'subheader'
            },
            {
                text: [
                    '\nWe enclose accounts listed hereunder for your tracing on the following terms and conditions:',
                    '\n',
                    '\nReceipt of all instructions issued by the Bank must be acknowledged in writing. Unless otherwise instructed by us in writing, INVESTIGATORS task is to undertake investigations to trace debtors’ whereabouts, ascertain their means and assets.',
                    '\n\nThe resultant report on the investigations (hereinafter referred to as the ‘report’) should be; ',
                    '\n•	Typed as per the template provided with illustrations/maps as necessary together with all documents/details on the debtors',
                    '\n•	Must be forwarded to us, within Thirty (30) days from date of this letter whether the investigation is completed or not.      ',
                    '\n\n(a)	 At all times during the currency of the investigations, investigator shall abstain from divulging any prejudicial information relating to our business and the debtors under your investigation. Accordingly, they should indemnify, release and discharge us from all actions, suits, claims, demands brought against us or yourselves solely and/or jointly out of disclosure of such injurious information. ',
                    '\n\n(b)	Please obtain and maintain at your cost an insurance cover with a reputable insurance company to cover all liability arising during the investigation. ',
                    '\n\n(c)	Either party may terminate the agreement for any reason by giving eight (8) days written notice to the other party.  ',
                    '\n\n(d)	Either of the party should refer all acute disputes, differences or questions on issues arising on the accounts under investigation for decision of an Arbitrator to be appointed in writing by the parties and if they cannot agree, to the Chairman of the Chartered Institute of Arbitrators, Kenya Chapter.',
                    '\n\n(e)	No investigator has authority whatsoever to negotiate repayment proposals of any kind or accept any form of monetary payment or enter into any other communication or correspondence with a debtor other than on the lines set out. Debtors must always be advised to address their proposals, queries or payments directly to the Bank.',
                    '\n\nIf you are agreeable to our terms, please sign below and commence investigations on the under listed account:',
                    '\n'
                ]
            },
            '\n',
            {
                alignment: 'justify',
                fontSize: 10,
                table: {
                    body: [
                        [{ text: '#', style: 'tableHeader', alignment: 'left' }, { text: 'Customer name', style: 'tableHeader', alignment: 'left' }, { text: 'RROCODE', style: 'tableHeader', alignment: 'left' }, { text: 'Reason for Investigation', style: 'tableHeader', alignment: 'left' }],
                        ['1', letter_data.custname, letter_data.rrocode, letter_data.reasonforinvestigation]
                    ]
                },
                layout: {
                    fillColor: function (rowIndex, node, columnIndex) {
                        return (rowIndex % 2 === 0) ? '#CCCCCC' : null;
                    }
                }
            },

      
            { text: '\n\n\nYours Faithfully,' },
            { text: '\n\nMITCHELLE ARUA,                                                                           Roseline Ogango,', style: 'tableHeader' },
            { text: 'Remedial Management Department                                         For Manager- Remedial Management Department', style: 'tableHeader' },
            { text: '\n\nThis letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },

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
    writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.accnumber + DATE + "investigators.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.accnumber + DATE + "investigators.pdf",
            filename: letter_data.accnumber + DATE + "investigators.pdf"
        })
    });
});

module.exports = router;
