const fs = require('fs')
const xlsx = require('xlsx')
const inquirer = require('inquirer')
const color = require('chalk')

const questions = [
  {
    type: 'input',
    name: 'type',
    message: '1. Que tipo de archivo desea procesar ? :  '
  },
  {
    type: 'input',
    name: 'fileName',
    message: '2. Ingrese el nombre del archivo a procesar :  '
  },
]

inquirer.prompt(questions)
  .then(answers => {
    processData(answers)
  })

const processData = (params) => {
  const { type, fileName } = params
  try {
    const workBook = xlsx.readFile(`./excel/${fileName}.xlsx`)
    const sheets = workBook.SheetNames.map(name => workBook.Sheets[name])

    const header = ['fecha','proveedor','co','nroFactura','nroRemision','bodega','referencia','color','talla','concat','cantidad','descCo','cotizacion','cotizacionTotal','codBarras']

    const data = sheets.map(sheet => {
      const json = xlsx.utils.sheet_to_json(sheet, { header, raw: false, dateNF: 'yyyymmdd' })
      json.shift()
      return json
    })

    const dataString = JSON.stringify(data, null, 2)

    fs.writeFile(`${fileName}.json`, dataString, 'utf8', (err) => {
      if(err) throw err

      console.log('\n')
      console.log(color.blue('file firlan.json was saved successfully!'))
    })
  } catch(err) {
    throw color.red(err.message)
  }
}
