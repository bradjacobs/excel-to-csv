/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;

class AdvancedExcelSheetReaderTest extends AbstractExcelSheetReaderTest<AdvancedExcelSheetReader, AdvancedExcelSheetReader.Builder> {

    @Override
    protected AdvancedExcelSheetReader.Builder createBuilder() {
        return AdvancedExcelSheetReader.builder();
    }

    // TODO
    //   1. add test when manually create new SharedStringsTable()
    //   2. add test when manually create new StylesTable()
}
