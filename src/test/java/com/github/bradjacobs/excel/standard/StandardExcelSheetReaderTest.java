/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;

public class StandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<StandardExcelSheetReader, StandardExcelSheetReader.Builder> {

    @Override
    protected StandardExcelSheetReader.Builder createBuilder() {
        return StandardExcelSheetReader.builder();
    }
}
