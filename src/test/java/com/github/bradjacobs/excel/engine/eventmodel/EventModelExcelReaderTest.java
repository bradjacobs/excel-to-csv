/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel;

import com.github.bradjacobs.excel.reader.AbstractExcelReaderTest;

class EventModelExcelReaderTest extends AbstractExcelReaderTest<EventModelExcelReader, EventModelExcelReader.Builder> {

    @Override
    protected EventModelExcelReader.Builder createBuilder() {
        return EventModelExcelReader.builder();
    }

    // TODO
    //   1. add test when manually create new SharedStringsTable()
    //   2. add test when manually create new StylesTable()
}
