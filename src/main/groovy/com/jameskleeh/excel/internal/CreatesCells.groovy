/*
Licensed to the Apache Software Foundation (ASF) under one
or more contributor license agreements.  See the NOTICE file
distributed with this work for additional information
regarding copyright ownership.  The ASF licenses this file
to you under the Apache License, Version 2.0 (the
        "License"); you may not use this file except in compliance
with the License.  You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing,
        software distributed under the License is distributed on an
"AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, either express or implied.  See the License for the
specific language governing permissions and limitations
under the License.
*/
package com.jameskleeh.excel.internal

import com.jameskleeh.excel.CellStyleBuilder
import com.jameskleeh.excel.Excel
import com.jameskleeh.excel.Formula
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * A base class used to create cells
 *
 * @author James Kleeh
 */
abstract class CreatesCells {

    protected final XSSFWorkbook workbook
    protected final XSSFSheet sheet
    protected Map defaultOptions
    protected final Map<Object, Integer> columnIndexes
    protected final CellStyleBuilder styleBuilder

    CreatesCells(XSSFWorkbook workbook, XSSFSheet sheet, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder) {
        this.workbook = workbook
        this.sheet = sheet
        this.defaultOptions = defaultOptions
        this.columnIndexes = columnIndexes
        this.styleBuilder = styleBuilder
    }

    /**
     * Sets the default styles to use for the given row or column
     *
     * @param options The style options
     */
    void defaultStyle(Map options) {
        options = new LinkedHashMap(options)
        styleBuilder.convertSimpleOptions(options)
        if (defaultOptions == null) {
            defaultOptions = options
        } else {
            defaultOptions = styleBuilder.merge(defaultOptions, options)
        }
    }

    protected abstract XSSFCell nextCell()

    /**
     * Skips cells
     *
     * @param num The number of cells to skip
     */
    abstract void skipCells(int num)

    protected void setStyle(Object value, XSSFCell cell, Map options) {
        styleBuilder.setStyle(value, cell, options, defaultOptions)
    }

    /**
     * Creates a header cell
     *
     * @param value The cell value
     * @param id The cell identifer
     * @param style The cell style
     * @return The native cell
     */
    XSSFCell column(String value, Object id, final Map style = null) {
        XSSFCell col = cell(value, style)
        columnIndexes[id] = col.columnIndex
        col
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param formulaString The formula
     * @param style The cell style
     * @return The native cell
     */
    XSSFCell formula(String formulaString, final Map style) {
        XSSFCell cell = nextCell()
        if (formulaString.startsWith('=')) {
            formulaString = formulaString[1..-1]
        }
        cell.setCellFormula(formulaString)
        setStyle(null, cell, style)
        cell
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param formulaString The formula
     * @return The native cell
     */
    XSSFCell formula(String formulaString) {
        formula(formulaString, null)
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param callable The return value will be assigned to the cell formula. The closure delegate contains helper methods to get references to other cells.
     * @return The native cell
     */
    XSSFCell formula(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Formula) Closure callable) {
        formula(null, callable)
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param style The cell style
     * @param callable The return value will be assigned to the cell formula. The closure delegate contains helper methods to get references to other cells.
     * @return The native cell
     */
    XSSFCell formula(final Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Formula) Closure callable) {
        XSSFCell cell = nextCell()
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new Formula(cell, columnIndexes)
        String formula
        if (callable.maximumNumberOfParameters == 1) {
            formula = (String)callable.call(cell)
        } else {
            formula = (String)callable.call()
        }
        if (formula.startsWith('=')) {
            formula = formula[1..-1]
        }
        cell.setCellFormula(formula)
        setStyle(null, cell, style)
        cell
    }

    /**
     * Creates a new blank cell
     *
     * @return The native cell
     */
    XSSFCell cell() {
        cell('')
    }

    /**
     * Creates a new cell and assigns a value
     *
     * @param value The value to assign
     * @return The native cell
     */
    XSSFCell cell(Object value) {
        cell(value, null)
    }

    /**
     * Creates a new cell with a value and style
     *
     * @param value The value to assign
     * @param style The cell style options
     * @return The native cell
     */
    XSSFCell cell(Object value, final Map style) {

        XSSFCell cell = nextCell()
        setStyle(value, cell, style)
        if (value instanceof String) {
            cell.setCellValue(value)
        } else if (value instanceof Calendar) {
            cell.setCellValue(value)
        } else if (value instanceof Date) {
            cell.setCellValue(value)
        } else if (value instanceof Number) {
            cell.setCellValue(value.doubleValue())
        } else if (value instanceof Boolean) {
            cell.setCellValue(value)
        } else {
            Closure callable = Excel.getRenderer(value.class)
            if (callable != null) {
                cell.setCellValue((String)callable.call(value))
            } else {
                cell.setCellValue(value.toString())
            }
        }
        cell
    }

    /**
     * Merges cells
     *
     * @param style Default styles for merged cells
     * @param callable To build cell data
     */
    abstract void merge(final Map style, Closure callable)

    /**
     * Merges cells
     *
     * @param callable To build cell data
     */
    abstract void merge(Closure callable)

    /**
     * Merges cells
     *
     * @param value The cell content
     * @param count How many cells to merge
     * @param style Styling for the merged cell
     */
    void merge(Object value, Integer count, final Map style = null) {
        merge(style) {
            cell(value)
            skipCells(count)
        }
    }

}