var express = require('express');
var router = express.Router();
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const cors = require('cors');
var Minio = require("minio");
const docx = require('docx');


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

const {
    TextRun,
    HorizontalPositionRelativeFrom,
    HorizontalPositionAlign
} = docx;
const { Document, Footer, Header, Packer, Paragraph, Table, TableCell, TableRow } = docx;


router.use(express.urlencoded({ extended: true }));
router.use(express.json());

router.use(cors());


router.get('/', function (req, res) {
    res.json({ message: 'Reposession send physically letter is ready!' });
});


router.post('/download', function (req, res) {
    const letter_data = req.body;
    const INCLUDELOGO = true;
    var date1 = new Date();
    const DATE = dateFormat(date1, "dd-mmm-yyyy");
    const rawaccnumber = letter_data.accnumber;
    const repodate = letter_data.daterepoissued;
    const repodate2 = dateFormat(repodate, "dd-mmm-yyyy");
    const first4 = rawaccnumber.substring(0, 9);
    const accnumber_masked = first4 + 'xxxxx';

    const doc = new Document({
        sections: [
            {
                headers: {
                    default: new Header({
                        children: [new Paragraph("Header text")],
                    }),
                },
                footers: {
                    default: new Footer({
                        children: [new Paragraph(data.footerfirst)],
                    }),
                },
                children: [new Paragraph("Hello World")],
            },
        ],
    });

    
    


    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync(LETTERS_DIR + accnumber_masked + DATE + "repossession.docx", buffer);
        // save to minio
        const filelocation = LETTERS_DIR + accnumber_masked + DATE + "repossession.docx";
        const bucket = 'demandletters';
        const savedfilename = accnumber_masked + '_' + Date.now() + '_' + "repossession.docx"
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
            }
            res.json({
                result: 'success',
                message: LETTERS_DIR + accnumber_masked + DATE + "repossession.docx",
                filename: accnumber_masked + DATE + "repossession.docx",
                savedfilename: savedfilename,
                objInfo: objInfo
            })
        });
        //save to mino end
    })




}); // end post
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
