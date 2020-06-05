const App = SpreadsheetApp
const sheet = App.getActiveSheet()
const regex: RegExp = /[Ａ-Ｚａ-ｚ０-９]/g

const convertFullToHalf = () => {
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()
  const targetCells: TargetCells = []

  const allCells = sheet.getRange(1, 1, lastRow, lastColumn).getValues()

  for (let i = 0; i < lastRow; i++) {
    for (let j = 0; j < lastColumn; j++) {
      const cell = String(allCells[i][j])
      const target = convertCharacters(cell)
      if (target !== null && target !== undefined && target !== '') {
        targetCells.push({
          value: target,
          row: i + 1,
          column: j + 1,
        })
      }
    }
  }
  for (let k in targetCells) {
    const row = targetCells[k].row
    const column = targetCells[k].column
    const value = targetCells[k].value
    sheet.getRange(row, column).setValue(value)
  }
}

type TargetCells = Cell[]
interface Cell {
  value: string
  row: number
  column: number
}

const convertCharacters = (string: string) => {
  if (string.match(regex) === null || string.slice(0, 1) === '=') {
    return
  }
  return string.replace(regex, (s) => {
    return String.fromCharCode(s.charCodeAt(0) - 0xfee0)
  })
}
