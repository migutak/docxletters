/*module.exports = {
    // imagePath: '/Users/kevinabongo/projects/docxletters/routes/',
    // imagePath: '/home/ecollectadmin/docxletters/routes/',
    filePath: 'C:\\demands\\',
    // filePath: "C:\\Users\\Kevin\\Documents\\angular2\\docxletters\\",
    imagePath: 'C:\\Users\\Kevin\\Documents\\angular2\\docxletters\\routes\\',
    pipeUrl: 'http://ecollectapp01.co-opbank.co.ke:4000/upload'
}*/

module.exports = {
    filePath        : process.env.FILEPATH || '/app/demands/', // "d:\\demands\\",
    imagePath       : process.env.IMAGEPATH || '/app/routes/' // 'd:\\angularprojects\\docxletters\\routes\\'
};