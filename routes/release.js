var express = require('express');
var router = express.Router();
const nodemailer = require("nodemailer");
const fs = require('fs');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const cors = require('cors');
var Minio = require("minio");
require('log-timestamp');

var minioClient = new Minio.Client({
    endPoint: process.env.MINIO_ENDPOINT || '127.0.0.1',
    port: process.env.MINIO_PORT ? parseInt(process.env.MINIO_PORT, 10) : 9005,
    useSSL: false,
    accessKey: process.env.ACCESSKEY || 'AKIAIOSFODNN7EXAMPLE',
    secretKey: process.env.SECRETKEY || 'wJalrXUtnFEMIK7MDENGbPxRfiCYEXAMPLEKEY'
});
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
    res.json({ message: 'Release letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    console.log(letter_data);
    const rawaccnumber = letter_data.acc;
    const first4 = rawaccnumber.substring(0, 9);
    accnumber_masked = first4 + 'xxxxx';

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
            '',
            '\n' + DATE,
            '\nThe Managing Director',
            '\n' + letter_data.storageyard,
            ' Head Office',

            '\nDear Sir/Madam',
            {
                text: '\nRe: RELEASE OF VEHICLE REG NO. ' + letter_data.vehicleregno ,
                style: 'subheader'
            },
            {
                text: [
                    '\nThe above captioned subject matter refers. ',
                    '\n',
                    '\n'
                ]
            },
            {
                text: [
                    'We confirm that the vehicle may be released to the Client’s representative  ',
                    { text: letter_data.custname, fontSize: 10, bold: true },
                    ' ID Number ',
                    { text: letter_data.nationalid, fontSize: 10, bold: true },
                    ' upon proper application, identification and settlement of auctioneer, and storage fees. ',
                    '\n',
                    '\n',
                    '\nYour co-operation is highly appreciated. The undersigned may be contacted on Tel no. ',
                    { text: letter_data.celnumber, fontSize: 10, bold: true },
                    ' in case of any queries. ',
                    '\n',
                    '\n',
                    'Yours faithfully, ',
                ]
            },
            '\n',
            {
                table: {
                    headerRows: 1,
                    widths: [200, '*'],
                    body: [
                      [{text: '', style: 'tableHeader'}, {text: '', style: 'tableHeader'}],
                      ['\n\n---------------------------, ', ' \n\n---------------------------------------'],
                      ['REMEDIAL OFFICER', 'MANAGER, ADMIN & SUPPORT REMEDIAL MGT.']
                    ]
                },
                layout: 'noBorders',
                
            },


            { text: '\nCc ' },
            { text: '' + letter_data.custname }
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
    writeStream = fs.createWriteStream(LETTERS_DIR + accnumber_masked + DATE + "release.pdf");
    pdfDoc.pipe(writeStream);
    pdfDoc.end();
    writeStream.on('finish', function () {
        // res.json({
        //     result: 'success',
        //     message: LETTERS_DIR + letter_data.accnumber + DATE + "release.pdf",
        //     filename: letter_data.accnumber + DATE + "release.pdf"
        // })
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "release.pdf";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "release.pdf"
        console.log(filelocation);
        console.log(savedfilename);
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
                deleteFile(filelocation);
            }
            res.json({
                result: 'success',
                message: LETTERS_DIR + accnumber_masked + DATE + "release.pdf",
                filename: accnumber_masked + DATE + "release.pdf",
                savedfilename: savedfilename,
                objInfo: objInfo
            })
            deleteFile(filelocation);
        });
        //save to mino end
    });
});
function deleteFile(req) {
    fs.unlink(req, (err) => {
        if (err) {
            console.error(err)
            return
        }
        //file removed
    })
}
module.exports = router;
