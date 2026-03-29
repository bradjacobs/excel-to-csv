/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;

class StandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<StandardExcelSheetReader, StandardExcelSheetReader.Builder> {

    @Override
    protected StandardExcelSheetReader.Builder createBuilder() {
        return StandardExcelSheetReader.builder();
    }
}
