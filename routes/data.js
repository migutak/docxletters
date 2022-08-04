module.exports = {
    // filePath        : process.env.FILEPATH || '/app/nfsmount/demandletters/', // "d:\\demands\\",
    filePath        : process.env.FILEPATH || "C:/demands/",
    imagePath       : process.env.IMAGEPATH || '/app/docxletters/routes/', // 'd:\\angularprojects\\docxletters\\routes\\',
    SENDEMAILURL     : process.env.SENDEMAILURL || 'http://172.16.204.72:8005/ipfcancellation/email',
    smtpport: process.env.SMTPPORT || 587,  // 465 FOR SECURE OR PROD
    smtpuser:  'allanmaroko10',
    pass:  process.env.PASS || 'Vipermarox411',
    smtpserver: process.env.SMTPSERVER || 'smtp.gmail.com',   // office365.officer
    secure: false,
    type: 'login',
    requireTLS: true,
    footerfirst : process.env.footerfirst || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), ',
    footersecond : process.env.footersecond || 'J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia',
    footeroneline : process.env.footersecond || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), J. Sitienei, B. Simiyu, P. Githendu, W.Ongoro, R.Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia'
};
