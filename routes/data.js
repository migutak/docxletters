module.exports = {
    filePath        : process.env.FILEPATH || "C:/Users/KevinMigutaAbongo/Documents/demands/", 
    imagePath       : process.env.IMAGEPATH || '/app/docxletters/routes/', // 'd:\\angularprojects\\docxletters\\routes\\',
    SENDEMAILURL     : process.env.SENDEMAILURL || 'http://172.16.204.72:8005/ipfcancellation/email',
    smtpserver: process.env.SMTPSERVER || 'smtp.gmail.com', //host: 'smtp.gmail.com',office365.officer
    smtpport: process.env.SMTPPORT || 587,
    smtpuser: process.env.SMTPUSER || 'ecollectsystem@gmail.com',
    pass:  process.env.PASS || 'W1ndowsxp',
    footerfirst : process.env.footerfirst || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), ',
    footersecond : process.env.footersecond || 'J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia',
    footeroneline : process.env.footersecond || 'Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group M.D & CEO), M. Malonza (Vice Chairman), J. Sitienei, B. Simiyu, P. Githendu, W.Ongoro, R.Kimanthi, W. Mwambia, W. Welton (Mrs), M. Karangatha (Mrs), L. Karissa, G. Mburia'
};
