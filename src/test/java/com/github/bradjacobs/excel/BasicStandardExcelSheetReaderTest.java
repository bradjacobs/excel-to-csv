package com.github.bradjacobs.excel;

public class BasicStandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<StandardExcelSheetReader, StandardExcelSheetReader.Builder> {

    @Override
    protected StandardExcelSheetReader.Builder createBuilder() {
        return StandardExcelSheetReader.builder();
    }
}
