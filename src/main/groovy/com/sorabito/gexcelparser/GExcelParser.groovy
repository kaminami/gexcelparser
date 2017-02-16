package com.sorabito.gexcelparser

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Sheet
import org.jggug.kobo.gexcelapi.CellLabelUtils as CLU
import org.jggug.kobo.gexcelapi.CellRange
import org.jggug.kobo.gexcelapi.GExcel


class GExcelParser {
    def book
    FormulaEvaluator formulaEvaluator


    static GExcelParser open(InputStream is) {
        def inst = new GExcelParser()
        inst.openBook(is)
        return inst
    }

    static GExcelParser open(File file) {
        def inst = new GExcelParser()
        inst.openBook(file)
        return inst
    }

    static GExcelParser open(String filePath) {
        def inst = new GExcelParser()
        inst.openBook(filePath)
        return inst
    }

    def openBook(InputStream is) {
        this.book = GExcel.open(is)
        this.prepare()
    }

    def openBook(File file) {
        this.book = GExcel.open(file)
        this.prepare()
    }

    def openBook(String filePath) {
        this.book = GExcel.open(filePath)
        this.prepare()
    }

    def prepare() {
        this.prepareFormulaEvaluatorFor()
    }

    def prepareFormulaEvaluatorFor() {
        this.formulaEvaluator = this.book.getCreationHelper().createFormulaEvaluator()
    }

    def parseSheet(String sheetName) {
        String stopLabel = this.findStopColumnLabel(sheetName)
        if (stopLabel == null) {
            throw new Exception('header not found')
        }

        return this.parseSheet(sheetName, 'A', stopLabel, 1, 2)
    }

    def findStopColumnLabel(String sheetName) {
        Sheet sheet = this.book."${sheetName}"


        def firstEmptyIndex = (1..16384).find { int idx ->
            def columnLabel = CLU.columnLabel(idx)
            def value = sheet."${columnLabel}1".value?.toString()

            (value == null) || value.isEmpty()
        }

        if (firstEmptyIndex == 1) { return null }
        return CLU.columnLabel(firstEmptyIndex - 1)
    }

    def parseSheet(String sheetName, String startColumnLabel, String stopColumnLabel, int headerRow, int firstDataRow) {
        Sheet sheet = this.book."${sheetName}"
        def columns = this.parseHeaderRow(sheet, startColumnLabel, stopColumnLabel, headerRow)


        def valueList = []
        int rowIndex = firstDataRow
        while(this.hasNextRow(sheet, startColumnLabel, rowIndex)) {
            valueList << this.parseDataRow(sheet, startColumnLabel, stopColumnLabel, rowIndex, columns)
            rowIndex++
        }

        return valueList
    }

    Map parseHeaderRow(Sheet sheet, String startColumnLabel, String stopColumnLabel, int headerRow) {
        CellRange cells = sheet."${startColumnLabel}${headerRow}_${stopColumnLabel}${headerRow}"

        Map columnMap = [:]
        cells.first().eachWithIndex { Cell cell, int idx ->
            def key = this.formulaEvaluator.evaluateInCell(cell).toString()
            columnMap[key] = idx
        }

        return columnMap
    }

    boolean hasNextRow(Sheet sheet, String startColumnLabel, int rowIndex) {
        def firstCell = sheet."${startColumnLabel}${rowIndex}"
        def value = this.formulaEvaluator.evaluateInCell(firstCell).value?.toString()

        if (value == null) { return false }
        if (value.isEmpty()) { return false }

        return true
    }

    PropertyMapper parseDataRow(Sheet sheet, String startColumnLabel, String stopColumnLabel, int rowIndex, Map columns) {
        CellRange cells = sheet."${startColumnLabel}${rowIndex}_${stopColumnLabel}${rowIndex}"

        def values = []

        cells.first().eachWithIndex { Cell cell, int idx ->
            values << this.formulaEvaluator.evaluateInCell(cell).value
        }

        def mapper = new PropertyMapper([columns:columns, values:values])
        return mapper
    }
}
