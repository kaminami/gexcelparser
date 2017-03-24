package com.sorabito.gexcelparser

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

import org.jggug.kobo.gexcelapi.CellLabelUtils as CLU
import org.jggug.kobo.gexcelapi.CellRange
import org.jggug.kobo.gexcelapi.GExcel


class GExcelParser {
    static final int MAX_COLUMNS = 16384 // Excel2007-

    Workbook book
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

    void openBook(InputStream is) {
        this.book = GExcel.open(is)
        this.prepare()
    }

    void openBook(File file) {
        this.book = GExcel.open(file)
        this.prepare()
    }

    void openBook(String filePath) {
        this.book = GExcel.open(filePath)
        this.prepare()
    }

    void prepare() {
        this.prepareFormulaEvaluatorFor()
    }

    void prepareFormulaEvaluatorFor() {
        this.formulaEvaluator = this.book.getCreationHelper().createFormulaEvaluator()
    }

    List<PropertyMapper> parseSheet(String sheetName) {
        String stopLabel = this.findStopColumnLabel(sheetName)
        if (stopLabel == null) {
            throw new Exception('header not found')
        }

        return this.parseSheet(sheetName, 'A', stopLabel, 1, 2)
    }

    String findStopColumnLabel(String sheetName) {
        Sheet sheet = this.book."${sheetName}"


        def firstEmptyIndex = (1 .. MAX_COLUMNS).find { int idx ->
            def columnLabel = CLU.columnLabel(idx)
            def value = sheet."${columnLabel}1".value?.toString()

            (value == null) || value.isEmpty()
        }

        if (firstEmptyIndex == 1) { return null }
        return CLU.columnLabel(firstEmptyIndex - 1)
    }

    List<PropertyMapper> parseSheet(String sheetName, String startColumnLabel, String stopColumnLabel, int headerRow, int firstDataRow) {
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
            values << this.extractCellValue(cell)
        }

        def mapper = new PropertyMapper([columns:columns, values:values])
        return mapper
    }

    def extractCellValue(Cell cell) {
        this.formulaEvaluator.evaluateInCell(cell).value


        switch(cell.getCellType()) {
            case Cell.CELL_TYPE_BLANK:
                return null

            case Cell.CELL_TYPE_STRING:
                return cell.getRichStringCellValue().getString()

            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue()

            case Cell.CELL_TYPE_NUMERIC:
                return this.extractNumericCellValue(cell)

            case Cell.CELL_TYPE_FORMULA:
                return this.formulaEvaluator.evaluateInCell(cell).value

            default:
                return null;
        }
    }

    def extractNumericCellValue(Cell cell) {
        // date
        if (DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue()
        }

        // double or long
        DataFormatter formatter = new DataFormatter()
        String retValue = formatter.formatCellValue(cell)

        if (retValue.contains('.')) { return Double.parseDouble(retValue) }
        return Long.parseLong(retValue)
    }
}
