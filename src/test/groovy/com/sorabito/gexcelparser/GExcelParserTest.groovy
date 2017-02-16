package com.sorabito.gexcelparser


class GExcelParserTest extends GroovyTestCase {

    void testParseSheet1() {
        def parser = GExcelParser.open('build/resources/test/sample.xlsx')
        List<PropertyMapper> mapperList = parser.parseSheet('Sheet1')

        assert mapperList.size() == 5


        // column name -> column index
        def columns = mapperList.first().columns

        assert columns.size() == 4
        assert columns['column1'] == 0
        assert columns['column2'] == 1
        assert columns['column3'] == 2
        assert columns['column4'] == 3


        // data row
        def firstRow = mapperList[0]
        assert firstRow['column1'] == 1
        assert firstRow['column2'] == 'aa'
        assert firstRow['column3'] == 1.1
        assert firstRow['column4'] == true

        def lastRow = mapperList[4]
        assert lastRow.'column1' == 5
        assert lastRow.'column2' == 'ee'
        assert lastRow.'column3' == 5.1
        assert lastRow.'column4' == false
    }

    void testParseSheet5() {
        def parser = GExcelParser.open('build/resources/test/sample.xlsx')
        List<PropertyMapper> mapperList = parser.parseSheet('Sheet2', 'C', 'F', 5, 7)

        assert mapperList.size() == 5


        // column name -> column index
        def columns = mapperList.first().columns

        assert columns.size() == 4
        assert columns['column1'] == 0
        assert columns['column2'] == 1
        assert columns['column3'] == 2
        assert columns['column4'] == 3


        // data row
        def firstRow = mapperList[0]
        assert firstRow['column1'] == 1
        assert firstRow['column2'] == 'aa'
        assert firstRow['column3'] == 1.1
        assert firstRow['column4'] == true

        def lastRow = mapperList[4]
        assert lastRow.'column1' == 5
        assert lastRow.'column2' == 'ee'
        assert lastRow.'column3' == 5.1
        assert lastRow.'column4' == false
    }
}
