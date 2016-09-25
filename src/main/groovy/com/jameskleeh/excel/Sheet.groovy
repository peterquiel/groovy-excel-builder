package com.jameskleeh.excel

import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A class used to create a sheet in an excel document
 */
@CompileStatic
class Sheet {

    private final XSSFSheet sheet
    private final XSSFWorkbook workbook
    private int rowIdx
    private Map defaultOptions
    private Map<Object, Integer> columnIndexes = [:]
    private final CellStyleBuilder styleBuilder

    private static final String HEIGHT = 'height'

    Sheet(XSSFWorkbook workbook, XSSFSheet sheet) {
        this.workbook = workbook
        this.sheet = sheet
        this.rowIdx = 0
        this.styleBuilder = new CellStyleBuilder(workbook)
    }

    void defaultStyle(Map options) {
        defaultOptions = options
        styleBuilder.convertSimpleOptions(defaultOptions)
    }

    void skipRows(int num) {
        rowIdx += num
    }

    void columns(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row(callable)
    }

    void columns(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row(options, callable)
    }

    void row() {
        row([:], null)
    }

    void row(Object...cells) {
        row {
            cells.eachWithIndex { val, idx ->
                cell(val)
            }
        }
    }

    void row(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        row([:], callable)
    }

    void row(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        XSSFRow row = sheet.createRow(rowIdx)
        if (options) {
            if (options.containsKey(HEIGHT)) {
                Object heightObj = options.remove(HEIGHT)

                if (heightObj instanceof Number) {
                    row.setHeightInPoints(((Number)heightObj).floatValue())
                }

            }
            styleBuilder.setStyle(row, options, defaultOptions)
        }
        if (callable != null) {
            callable.resolveStrategy = Closure.DELEGATE_FIRST
            callable.delegate = new Row(workbook, sheet, row, defaultOptions, columnIndexes)
            if (callable.maximumNumberOfParameters == 1) {
                callable.call(row)
            } else {
                callable.call()
            }
        }
        rowIdx++
    }
}
