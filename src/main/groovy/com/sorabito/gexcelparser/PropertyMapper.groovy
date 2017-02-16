package com.sorabito.gexcelparser


class PropertyMapper {
    def values = []
    def columns = [:]

    def propertyMissing(String name) {
        def index = columns[name]
        if (index != null) {
            values[index]
        } else {
            throw new MissingPropertyException(name)
        }
    }

    def getAt(Integer index) {
        values[index]
    }

    String toString() {
        columns.collect { key, index -> "$key: ${values[index]}" }.join(', ')
    }

    Map toMap() {
        def sortedKeys = columns.keySet().sort { columns[it] }
        [sortedKeys, values].transpose().collectEntries()
    }
}
