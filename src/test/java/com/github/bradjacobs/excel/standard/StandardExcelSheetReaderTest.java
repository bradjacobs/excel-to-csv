/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.net.URL;
import java.nio.file.Path;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileUrl;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

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
