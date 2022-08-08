module.exports = {
    // filePath        : process.env.FILEPATH || '/app/nfsmount/demandletters/', // "d:\\demands\\",
    filePath        : process.env.FILEPATH || "C:/demands/",
    imagePath       : process.env.IMAGEPATH || '/app/docxletters/routes/', // 'd:\\angularprojects\\docxletters\\routes\\',
    SENDEMAILURL     : process.env.SENDEMAILURL || 'http://172.16.204.72:8005/ipfcancellation/email',
    smtpport: process.env.SMTPPORT || 2525,  // 465 FOR SECURE OR PROD
    smtpuser:  process.env.SMTPUSER || 'b6eef9a1d905d6',
    pass:  process.env.PASS || 'e4a461f71a936a',
    smtpserver: process.env.SMTPSERVER || 'smtp.mailtrap.io',   // office365.officer
    secure: false,
    type: 'login',
    requireTLS: true,
    footerfirst : process.env.footerfirst || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), ',
    footersecond : process.env.footersecond || 'J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia',
    footeroneline : process.env.footersecond || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), J. Sitienei, B. Simiyu, P. Githendu, W.Ongoro, R.Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia'
};
