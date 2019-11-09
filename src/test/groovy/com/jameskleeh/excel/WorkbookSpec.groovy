package com.jameskleeh.excel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import spock.lang.Specification

class WorkbookSpec extends Specification {

    void "test sheet"() {
        XSSFWorkbook workbook = ExcelBuilder.build {
            sheet {

            }
            sheet('Other') {

            }
            sheet([height: 12F, width: 20]) {

            }
        }

        when:
        Sheet other = workbook.getSheetAt(1)
        Sheet config = workbook.getSheetAt(2)

        then:
        workbook.numberOfSheets == 3
        other.sheetName == 'Other'
        config.defaultRowHeightInPoints == 12F
        config.defaultColumnWidth == 20
    }

    void "test open with template and use sheet by name access"() {
        given:
        XSSFWorkbook workbook = ExcelBuilder.build(this.class.getResourceAsStream('/template.xlsx')) {
            sheet('sheet') {
                row {
                    cell('created with builder')
                }
            }
        }

        when:
        XSSFSheet sheet = workbook.getSheet('sheet')
        Iterator<Cell> buildRow = sheet.getRow(0).cellIterator()
        Iterator<Cell> rowFromTemplate = sheet.getRow(1).cellIterator()

        then:
        buildRow.next().stringCellValue == 'created with builder'
        rowFromTemplate.next().stringCellValue == 'from template'
    }

}
