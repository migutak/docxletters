module.exports = {
    // filePath        : process.env.FILEPATH || '/app/nfsmount/demandletters/', // "d:\\demands\\",
    filePath        : process.env.FILEPATH || "C:\\Users\\pap\\Documents\\demands",
    imagePath       : process.env.IMAGEPATH || '/app/docxletters/routes/', // 'd:\\angularprojects\\docxletters\\routes\\',
    SENDEMAILURL     : process.env.SENDEMAILURL || 'http://172.16.204.72:8005/ipfcancellation/email',
    smtpserver: process.env.SMTPSERVER || 'smtp.gmail.com', //host: 'smtp.gmail.com',office365.officer
    smtpport: process.env.SMTPPORT || 587,
    smtpuser: process.env.SMTPUSER || 'ecollectsystem@gmail.com',
    pass:  process.env.PASS || 'W1ndowsxp',
};
