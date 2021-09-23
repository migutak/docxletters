module.exports = {
    // filePath        : process.env.FILEPATH || '/app/nfsmount/demandletters/', // "d:\\demands\\",
    filePath        : process.env.FILEPATH || "C:\\Users\\allan\\Videos\\generatedletters\\",
    imagePath       : process.env.IMAGEPATH || '/app/docxletters/routes/', // 'd:\\angularprojects\\docxletters\\routes\\',
    SENDEMAILURL     : process.env.SENDEMAILURL || 'http://172.16.204.72:8005/ipfcancellation/email',
    footerfirst : process.env.footerfirst || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), ',
    footersecond : process.env.footersecond || 'J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia'
};