//get excel's data in json 
const excelToObj = require('convert-excel-to-json');
// fs - file system
const fs = require('fs');
//read directory quizes
const files = fs.readdirSync('../quizes');
// console.log('test: ', files);

const createDir = (dirName) => {
    try {
        if (!fs.existsSync(`../parsedQuizes/${dirName}`)) {
            fs.mkdirSync(`../parsedQuizes/${dirName}`, 0777);
        }
    } catch (err) {
        console.error(err);
    }
};

const createFile = (dirName, list, text) => {
    try {
        fs.writeFileSync(`../parsedQuizes/${dirName}/${list}.txt`, text.join('\n\n'), {encoding:'utf8',flag:'w'});
    } catch (err) {
        console.error("Error with import text to file");
    }
};

const getContent = (lists, result, dirName) => {
    lists.map(list => {
        const text = result[list].map((objRow, rowI) => {
            const {
                B: question, 
                D: correctAnswer,
                E: opt1,
                F: opt2,
                G: opt3,
                H: opt4 } = objRow;
            
            const lettersDict = {
                0: (i) => `${i + 1}.`,
                1: () => 'a)',
                2: () => 'b)',
                3: () => 'c)',
                4: () => 'd)'
            };
            const mustntConsist = ['Question', 'Option 1', 'Option 2', 'Option 3', 'Option 4']
            const content = [question, opt1, opt2, opt3, opt4]
                    .map((el, i) => {
                        if (el === undefined || mustntConsist.includes(el)) {
                            rowI = 1;
                            return '';
                        }
                        if (i === correctAnswer)  {
                            
                            return `*${lettersDict[i]()} ${el}`
                        } else {
                            return `${lettersDict[i](rowI)} ${el}`
                        } 
                    })
                    .filter((el) => el !== '');
            return content.join('\n');
        });
        createFile(dirName, list, text);
        // try {
        //     fs.writeFileSync(`../${fileName}/${list}.txt`, text.join('\n'), {encoding:'utf8',flag:'w'});
        // } catch (err) {
        //     // console.error("Error with import text to file");
        // }
    });
    
};

files.forEach((file) => {
    const [dirName] = file.split('.');

    createDir(dirName);

    const result = excelToObj({
        sourceFile: `../quizes/${file}`,
        header: {
            rows: 1
        }
    });

    const listSheets = Object.keys(result);
    getContent(listSheets, result, dirName);
});
