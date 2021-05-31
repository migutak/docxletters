module.exports = {
    // filePath        : process.env.FILEPATH || '/app/nfsmount/demandletters/', // "d:\\demands\\",
    filePath        : process.env.FILEPATH || "d:\\demands\\",
    imagePath       : process.env.IMAGEPATH || '/app/docxletters/routes/', // 'd:\\angularprojects\\docxletters\\routes\\',
    SENDEMAILURL     : process.env.SENDEMAILURL || 'http://172.16.204.72:8005/ipfcancellation/email'
};
