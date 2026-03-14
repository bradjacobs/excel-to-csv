/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

public class BasicStandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<StandardExcelSheetReader, StandardExcelSheetReader.Builder> {

    @Override
    protected StandardExcelSheetReader.Builder createBuilder() {
        return StandardExcelSheetReader.builder();
    }
}
