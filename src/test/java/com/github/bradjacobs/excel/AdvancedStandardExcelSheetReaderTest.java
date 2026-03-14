package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.advanced.AdvancedExcelSheetReader;

public class AdvancedStandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<AdvancedExcelSheetReader, AdvancedExcelSheetReader.Builder> {

    @Override
    protected AdvancedExcelSheetReader.Builder createBuilder() {
        return AdvancedExcelSheetReader.builder();
    }
}
