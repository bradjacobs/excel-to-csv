/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static org.junit.jupiter.api.Assertions.assertThrows;

class StandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<StandardExcelSheetReader, StandardExcelSheetReader.Builder> {

    @Override
    protected StandardExcelSheetReader.Builder createBuilder() {
        return StandardExcelSheetReader.builder();
    }

    private static final Path VALID_TEST_INPUT_PSWD_PATH = getResourceFilePath("test_data_w_pswd_1234.xlsx");

    @Test
    public void nullSheetParameter() throws Exception {
        StandardExcelSheetReader sheetReader = StandardExcelSheetReader.builder().build();
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            sheetReader.convertToSheetContentArray(null);
        });
    }
}
