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

    var DATEEXPIRY = dateFormat(date1.setDate(date1.getDate() + 30), "dd-mmm-yyyy");
    var docDefinition = {
        pageSize: 'A4',
        pageOrientation: 'portrait',
        pageMargins: [50, 60, 50, 60],
        footer: {
            columns: [
                { text: 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group Managing Director & CEO), M. Malonza (Vice Chairman),J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, R. Simani (Mrs), L. Karissa, G. Mburia.\n\n' }
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
                            'Co-operative Bank House',
                            'Haile Selassie Avenue',
                            'P.O. Box 48231-00100 GPO, Nairobi',
                            'Tel: (020) 3276100',
                            'Fax: (020) 2227747/2219831',
                            { text: 'www.co-opbank.co.ke', color: 'blue', link: 'http://www.co-opbank.co.ke' }
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
                      ['To: Auctioneers name', 'Valid up to: ' + DATEEXPIRY],
                      ['Nairobi', ''],
                      ['', ''],
                      ['Asset Finance Agreement No.', letter_data.customerNumber],
                      ['', ''],
                      ['Hirer’s Name ', ':     ' + letter_data.ipfBalance],
                      ['Unit Financed ', ':     ' + letter_data.vehicleMake + ' & ' + letter_data.vehicleModel],
                      ['Registration No ', ':     ' + letter_data.vehicleRegnumber],
                    ]
                },
                layout: 'noBorders',
                
            },
            '\n',
            
            {
                text: [
                    '\nAccording to our records, the monthly rental of the above Asset finance Agreement is now in arrears. The total amount due is ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.ipfBalance)).format('0,0.00'), fontSize: 10, bold: true },
                    ' & ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.afBalance)).format('0,0.00'), fontSize: 10, bold: true },
                    ' and ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.odBalance)).format('0,0.00'), fontSize: 10, bold: true },
                    ' and others.',
                ]
                
            },
            {
                text: [
                    'A further rental of ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.nextScheduleAmount)).format('0,0.00'), fontSize: 10, bold: true },
                    ' becomes due on ',
                    { text: letter_data.nextScheduleDate, fontSize: 10, bold: true },
                    ' cumulatively amounting to ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.totalAmount)).format('0,0.00'), fontSize: 10, bold: true },
                    '.',
                ]
            },
            {
                text: [
                    '\nPlease approach the above named Hirer on our behalf and collect the total sum of ',
                    { text: 'Kes. ' + numeral(Math.abs(letter_data.totalAmount)).format('0,0.00'), fontSize: 10, bold: true },
                    ' plus your own charges, failing which you may take this letter as your authority to effect immediate re-possession of the above/equipment without further reference to us. HIRER MUST MAKE PAYMENT VIDE CASH OR BY BANKER’S CHEQUE AS PERSONAL CHEQUE(S) WILL NOT BE ACCEPTED. From our records, we are able to give the following additional information regarding this Agreement, which may assist you in your task of locating the hirer and/or the motor vehicle/equipment:-',
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
                      ['Telephone', ':     ' + letter_data.phoneNumber || 'N/A'],
                      ['Physical Address/Location', ':     ' + letter_data.place || 'N/A'],
                      ['Type of Business', ':      ' + letter_data.typeofbusiness || 'N/A'],
                      ['Bankers and Branch', ':     ' + letter_data.branchName || 'N/A'],
                      ['Purpose of Vehicle', ':     ' + letter_data.purposeofVehicle || 'N/A'],
                      ['Guarantors', ':     ' + letter_data.Guarantors || 'N/A'],
                      ['Guarantors Address', ':     ' + letter_data.GuarantorsAddress || 'N/A'],
                      ['Chassis No.', ':     ' + letter_data.chasisNumber || 'N/A'],
                      ['Engine No.', ':     ' + letter_data.EngineNo || 'N/A'],
                      ['Any other information', ':     ' + letter_data.otherInformation || 'N/A']
                    ]
                },
                layout: 'noBorders',
                
            },

            { text: '\n\nVehicle tracked by: ' + letter_data.trackingCompany, fontSize: 10, alignment: 'left' },

            { text: '\nYours Faithfully,' },
            { text: '\n\nAUTHORISED SIGNATORY,                                                                           AUTHORISED SIGNATORY', style: 'tableHeader' },
            { text: 'This letter is electronically generated and is valid without a signature ', fontSize: 9, italics: true, bold: true },


            { text: '\nCc ' },
            { text: '' + letter_data.customerName },
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
    writeStream = fs.createWriteStream(LETTERS_DIR + letter_data.accnumber + DATE + "repossession.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        res.json({
            result: 'success',
            message: LETTERS_DIR + letter_data.accnumber + DATE + "repossession.pdf",
            filename: letter_data.accnumber + DATE + "repossession.pdf"
        })
    });
});

module.exports = router;
