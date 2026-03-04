package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.advanced.AdvancedExcelSheetReader;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class AdvancedExcelSheetReaderTest extends AbstractExcelSheetDataExtractorTest<AdvancedExcelSheetReader, AdvancedExcelSheetReader.Builder> {

    @Override
    protected AdvancedExcelSheetReader.Builder createBuilder() {
        return AdvancedExcelSheetReader.builder();
    }
}
