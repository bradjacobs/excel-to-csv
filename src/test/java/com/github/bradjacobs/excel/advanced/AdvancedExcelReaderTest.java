/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.core.AbstractExcelReaderTest;

class AdvancedExcelReaderTest extends AbstractExcelReaderTest<AdvancedExcelReader, AdvancedExcelReader.Builder> {

    @Override
    protected AdvancedExcelReader.Builder createBuilder() {
        return AdvancedExcelReader.builder();
    }

    // TODO
    //   1. add test when manually create new SharedStringsTable()
    //   2. add test when manually create new StylesTable()
}
