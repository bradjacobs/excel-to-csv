package com.github.bradjacobs.excel;

public class BasicExcelSheetReaderTest extends AbstractExcelSheetDataExtractorTest<ExcelSheetReader, ExcelSheetReader.Builder> {

    @Override
    protected ExcelSheetReader.Builder createBuilder() {
        return ExcelSheetReader.builder();
    }
}
