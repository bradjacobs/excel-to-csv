/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;

public class AdvancedExcelSheetReaderTest extends AbstractExcelSheetReaderTest<AdvancedExcelSheetReader, AdvancedExcelSheetReader.Builder> {

    @Override
    protected AdvancedExcelSheetReader.Builder createBuilder() {
        return AdvancedExcelSheetReader.builder();
    }
}
