const fs = require('fs');
const path = require('path');
const inquirer = require('inquirer');
const hwp = require('node-hwp');
const extract = require('pdf-text-extract')
const parseString = require('xml2js').parseString;
const { initialize } = require('koalanlp/Util');
const { Tagger } = require('koalanlp/proc');
const { EUNJEON } = require('koalanlp/API');

const TARGET_DIR = path.resolve(__dirname, './targets');
const SAVE_DIR = path.resolve(__dirname, './results');

const files = fs.readdirSync('./targets');
let selectedFileName = '';

if (!fs.existsSync(TARGET_DIR)) {
  fs.mkdirSync(TARGET_DIR, { recursive: true })
}

if (!fs.existsSync(SAVE_DIR)) {
  fs.mkdirSync(SAVE_DIR, { recursive: true })
}

inquirer.prompt([{
  type: 'rawlist',
  name: 'target',
  message: '파일 확장자 선택',
  choices: ['hwp', 'pdf']
}]).then((choice) => {

  var targetFiles = files.filter(function(file) {
    return path.extname(file).toLowerCase() === `.${choice.target}`;
  });

  if (targetFiles.length > 0) {
    inquirer.prompt([{
      type: 'rawlist',
      name: 'target',
      message: '파일 선택',
      choices: targetFiles
    }]).then((choice) => {
      const fileName = choice.target;

      // extension 제거. filename만 추출.
      selectedFileName = fileName.split('.').slice(0, -1).join('.');

      switch(path.extname(fileName)) {
        case '.hwp':
          extractTextHWP(TARGET_DIR + '/' + fileName);
          break;
        case '.pdf':
          extractTextPDF(TARGET_DIR + '/' + fileName)
          break;
      }
    });
  } else {
    console.log(`선택 가능한 (${choice.target})파일이 없습니다. '${TARGET_DIR}' 디렉토리에 추출할 ${choice.target} 파일을 넣어주세요.`);
  }
});

const extractTextPDF = (filePath) => {
  extract(filePath, { splitPages: false }, function (err, pages) {
    if (err) {
      console.error(err);
      return
    }

    tokenizeString(pages.join(' '), 'pdf').then(
      () => console.log('단어 사전 추출이 완료되었습니다.'),
      (error) => console.error('Error occurred!', error)
    );;
  })
}

const extractTextHWP = (filePath) => {

  hwp.open(filePath, function(err, doc){
    var hml = doc.toHML();

    fs.writeFile(path.resolve(__dirname + '/results/test.xml'), hml, function(err) {
      var xml = fs.readFileSync('./results/test.xml', 'utf-8');
      var contents = '';
      parseString(xml, function(err, result) {
        // const title = result.HWPML.HEAD.DOCSUMMARY.TITLE;
        const body = result.HWPML.BODY;

        for (var i = 0, bMax = body.length; i < bMax; i += 1) {
          var section = body[i].SECTION;

          for (var j = 0, smax = section.length; j < smax; j += 1) {
            var p = section[i].P;
            var content = extractTextFromP(p);

            contents += content + ' ';
          }
        }

        tokenizeString(contents, 'hwp').then(
          () => console.log('단어 사전 추출이 완료되었습니다.'),
          (error) => console.error('Error occurred!', error)
        );
      });
    });

  });
}

const extractTextFromP = (pml) => {
  var p = pml;
  var string = '';

  for (var k = 0, pmax = p.length; k < pmax; k += 1) {
    var textList = p[k].TEXT;

    if (textList) {
      if (textList[0].CHAR) {
        var value = textList[0].CHAR[0];
        if (typeof value === 'string') {
          string += value + '\n';
        } else {
          if (value._) {
            string += value._;
          } else {
            console.log('misssing1 : ', value);
          }
        }

      }

      if (textList[0].TABLE && textList[0].TABLE[0].ROW) {
        for (var i = 0, max = textList[0].TABLE[0].ROW.length ; i < max; i += 1) {
          var cell = textList[0].TABLE[0].ROW[i].CELL;

          for (var c = 0, cmax = cell.length; c < cmax; c += 1) {
            if (cell[c].PARALIST[0].P) {
              if (typeof extractTextFromP(cell[c].PARALIST[0].P) === 'string') {
                string += extractTextFromP(cell[c].PARALIST[0].P) + ' ';
              } else {
                console.log('misssing2 : ', extractTextFromP(cell[c].PARALIST[0].P));
              }

            }
          }
        }
      }
    }
  }

  return string;
}

const rankingSorter = (firstKey, secondKey) => {
  return function(a, b) {
    if (a[firstKey] > b[firstKey]) {
      return -1;
    } else if (a[firstKey] < b[firstKey]) {
      return 1;
    }
    else {
      if (a[secondKey] > b[secondKey]) {
        return 1;
      } else if (a[secondKey] < b[secondKey]) {
        return -1;
      } else {
        return 0;
      }
    }
  }
}

async function tokenizeString(contents, extension){
  // ....
  await initialize({packages: {EUNJEON: '2.1.6', KKMA: '2.0.4'}, verbose: false});
  let tagger = new Tagger(EUNJEON);

  let results = await tagger(contents);
  let listMap = {};

  results.forEach((result) => {
    result.forEach((syntacticWord) => {
      syntacticWord.forEach((word) => {
        /**
         * NNG 보통 명사
         * NNP 고유 명사
         * NNB 일반 의존 명사
         * NNM 단위 의존 명사
         * */
        if (['NNG', 'NNP'].indexOf(word.getTag().tagname) !== -1 && word.getSurface().length > 1) {
          if (listMap[word.getSurface()]) {
            listMap[word.getSurface()]++;
          } else {
            listMap[word.getSurface()] = 1;
          }
        }
      })
    })
  })

  var listMapKeys = Object.keys(listMap);
  var list = [];
  var output = '';

  for (var i = 0, max = listMapKeys.length; i < max; i += 1) {
    var name = listMapKeys[i];
    var score = listMap[name];

    list.push({
      name, score
    });
  }

  const sorted = list.sort(rankingSorter("score", "name"));

  sorted.forEach((rank) => {
    output += `${rank.name} : ${rank.score} \r\n`;
  })

  console.log("extracted word count : " + sorted.length);

  fs.writeFile(path.resolve( `${SAVE_DIR}/${selectedFileName}-${extension}-extracted-content.txt`), contents, (err) => {
    console.log(`created ${SAVE_DIR}/${selectedFileName}-extracted-content.txt`);
  })

  fs.writeFile(path.resolve( `${SAVE_DIR}/${selectedFileName}-${extension}-word-list.txt`), output, (err) => {
    console.log(`created ${SAVE_DIR}/${selectedFileName}-word-list.txt`);
  })
}
