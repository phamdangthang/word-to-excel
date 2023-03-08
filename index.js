const reader = require('xlsx')
const WordExtractor = require("word-extractor")
const fs = require('fs')

const wordPath = __dirname+`/data/`
const excelPath = __dirname+`/export/`

fs.readdirSync(wordPath).forEach(wordFile => {
  console.log('Converting: ', wordFile);
  handle(wordFile);
}); 

function handle(wordFile) {
  const filePath = wordPath + wordFile
  const fileName = wordFile.split('.')['0']
  const excelFile = excelPath + fileName + '.xlsx'

  const extractor = new WordExtractor()
  const extracted = extractor.extract(filePath)

  fs.writeFile(excelFile, '', function (err) {
    if (err) {
      console.log('Failed! Close the file before doing it')
    }
  })

  extracted.then(function(doc) {
    const contents = doc.getBody().replace(/\t/g, '\n').split('\n').filter(function (el) {
      return el
    })

    const result = filteredData(contents)
    const ws = reader.utils.json_to_sheet(result)
    const file = reader.readFile(excelFile)
    reader.utils.book_append_sheet(file, ws)
    reader.writeFile(file, excelFile)
    console.log('Converted: ', fileName + '.xlsx')
  })
}

function filteredData(inputs) {
  const filtered = []
  const output = []
  let group = []
  inputs.forEach(item => {
    if (item.match(/^[0-9]/)) {
      if (group.length > 0) {
        filtered.push(group)
        group = []
      }
    }

    group.push(item)
  })
  filtered.push(group)

  for (const arr of filtered) {
    const obj = {'Nội dung': arr[0]}
    for (let i = 1; i < arr.length; i++) {
      obj[`Đáp án ${String.fromCharCode(64 + i)}`] = arr[i]
    }
    output.push(obj)
  }

  return output
}