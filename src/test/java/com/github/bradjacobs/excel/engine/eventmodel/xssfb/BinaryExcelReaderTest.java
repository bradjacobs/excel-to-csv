/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import com.github.bradjacobs.excel.engine.eventmodel.EventModelExcelReader;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.IOException;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class BinaryExcelReaderTest {

    private static final String HIDDEN_CELLS_DATA_FILE = "skipHiddenTestData.xlsb";
    private final Path HIDDEN_CELLS_FILE = TestResourceUtil.getResourceFilePath(HIDDEN_CELLS_DATA_FILE);
    private static final String INPUT_DATA_SHEET_SUFFIX = "_Data";
    private static final String EXPECTED_DATA_SHEET_SUFFIX = "_Expected";

    private final EventModelExcelReader defaultSheetReader = EventModelExcelReader.builder().build();


    @ParameterizedTest(name = "HiddenTest {index}: Sheet = {0}")
    @ValueSource(strings = {
            "BaseCase",
            "LastHidden",
            "MiddleBlankColumns",
            "MiddleBlankRows",
            "SingleHIddenRow",
            "SingleHIddenColumn",
            "LastColumn",
            "FirstLastRow",
            "AllColumnHidden",
            "FirstLastColumn",
            "Multi",
            "LastValueHidden",
            "LongestRowHidden"
    })
    void testMissingRowsAndColumns(String sheetNamePrefix) throws IOException {
        String testDataSheetName = sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX;
        String expectedDataSheetName = sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX;

        EventModelExcelReader sheetReader = EventModelExcelReader.builder()
                .skipHiddenCells(true)
                .build();

        ExcelReadRequest request = createRequest(HIDDEN_CELLS_FILE, testDataSheetName);
        ExcelReadRequest expectedRequest = createRequest(HIDDEN_CELLS_FILE, expectedDataSheetName);

        String[][] actualMatrix = readMatrix(sheetReader, request);
        if (actualMatrix.length > 0) {
            assertTrue(actualMatrix[0].length > 0, "Matrix return rows with zero-length arrays");
        }

        String[][] expectedMatrix = readDefaultMatrix(expectedRequest);
        assertArrayEquals(expectedMatrix, actualMatrix);
    }

    private ExcelReadRequest createRequest(Path path, String sheetName) {
        return ExcelReadRequest.from(path).bySheetName(sheetName).build();
    }

    private ExcelReadRequest createRequest(Path path, int sheetIndex) {
        return ExcelReadRequest.from(path).bySheetIndex(sheetIndex).build();
    }

    private String[][] readMatrix(EventModelExcelReader sheetReader, ExcelReadRequest request) throws IOException {
        return sheetReader.readSheet(request).getMatrix();
    }

    private String[][] readDefaultMatrix(ExcelReadRequest request) throws IOException {
        return readMatrix(defaultSheetReader, request);
    }
}
