/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.net.URL;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ExcelSkipHiddenDataTest {

    private static final String TEST_DATA_FILE = "skipHiddenTestData.xlsx";
    private static final String INPUT_DATA_SHEET_SUFFIX = "_Data";
    private static final String EXPECTED_DATA_SHEET_SUFFIX = "_Expected";

    // side note: the 'allColumnHidden' scenario will return empty rows,
    //   which can be misleading.  Should fix

    @ParameterizedTest(name = "HiddenTest {index}: Sheet = {0}")
    @ValueSource(strings = {"BaseCase", "LastColumn", "FirstLastRow",
            "AllColumnHidden", "FirstLastColumn", "Multi"})
    public void testMissingRowsAndColumns(String sheetNamePrefix) throws Exception {
        URL inputFile = getTestResourceFileUrl(TEST_DATA_FILE);

        ExcelReader dataExcelReader = ExcelReader.builder()
                .sheetName(sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX)
                .skipInvisibleCells(true)
                .build();
        String[][] actualMatrix = dataExcelReader.convertToDataMatrix(inputFile);

        ExcelReader expectedExcelReader = ExcelReader.builder()
                .sheetName(sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX)
                .build();
        String[][] expectedMatrix = expectedExcelReader.convertToDataMatrix(inputFile);

        assertArrayEquals(expectedMatrix, actualMatrix);
    }

    private URL getTestResourceFileUrl(String fileName) {
        URL resourceUrl = this.getClass().getClassLoader().getResource(fileName);
        assertNotNull(resourceUrl);
        return resourceUrl;
    }
}
