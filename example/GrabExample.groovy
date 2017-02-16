@GrabResolver(name="bintray-gexcelapi", root="http://dl.bintray.com/nobeans/maven/")
@GrabResolver(name="github-gexcelparser", root="https://github.com/sorabito/gexcelparser/raw/master/repository/")
@Grab("com.sorabito:gexcelparser:1.0.0")

import com.sorabito.gexcelparser.GExcelParser
import com.sorabito.gexcelparser.PropertyMapper

def parser = GExcelParser.open('example.xlsx')
List<PropertyMapper> mapperList = parser.parseSheet('Sheet1')

mapperList.each { mapper ->
    println "${mapper.column1}: ${mapper.column2}"
}

