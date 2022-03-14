import Excel from 'exceljs'
import { join } from 'path'
import { createReadStream, createWriteStream, existsSync } from 'fs'

console.log('==========================================================')
console.log('셀 병합 작업 시작.')

const filename = process.argv[2]
if (!filename) throw Error('파일명을 입력해주세요.')

const path = join('files', filename)
if (!existsSync(path)) throw Error(`파일이 존재하지 않습니다. 탐색경로 : ${path}`)

const sheetname = process.argv[3]
if (!sheetname) throw Error('시트명을 입력해주세요.')

// Workbook instance
const workbook = new Excel.Workbook()

// Initalize the workbook from file
await workbook.xlsx.read(createReadStream(path))

// Fetch the sheet
const sheet = workbook.getWorksheet(sheetname)
if (!sheet) throw Error('시트가 존재하지 않습니다.')

// Merge vertically
let fromRow = 3 // start row number = 3
let toRow = 4

while (true) {
    if (!sheet.getCell(`B${fromRow}`).value) break // 더이상 텍스트가 없음 (테이블 끝)

    while (true) {
        const text = sheet.getCell(`B${toRow}`).value
        if (text && text.includes('요약')) {
            horizontalMerge(toRow)
            toRow -= 1
            break
        }
        toRow += 1
        if (toRow >= 10_000)
            throw new Error(`요약 데이터를 찾을수 없습니다. 원재료명 : ${sheet.getCell(`B${fromRow}`).value}`)
    }

    sheet.mergeCells(`B${fromRow}:B${toRow}`) // 요약이 있는 cell의 앞 cell 까지 병합

    fromRow = toRow + 2
    toRow = toRow + 3
}

// Merge horizontally
function horizontalMerge(rowNum) {
    sheet.mergeCells(`B${rowNum}:F${rowNum}`)
}

// Write file
workbook.removeWorksheet(workbook.worksheets[0].id)
const newFileName = `${filename.replace('.xlsx', '')}_merged.xlsx`
const newPath = join('files', newFileName)
await workbook.xlsx.write(createWriteStream(newPath))

console.log('셀 병합 작업 완료.')
console.log(`저장된 경로:${newPath}`)
console.log('==========================================================')