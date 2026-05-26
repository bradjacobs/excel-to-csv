package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.engine.eventmodel.AdvancedExcelReader;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.util.TestResourceUtil;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.IOException;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class BinaryExcelReaderTest {

    private static final String HIDDEN_CELLS_DATA_FILE = "skipHiddenTestData.xlsb";
    private final Path HIDDEN_CELLS_FILE = TestResourceUtil.getResourceFilePath(HIDDEN_CELLS_DATA_FILE);
    private static final String INPUT_DATA_SHEET_SUFFIX = "_Data";
    private static final String EXPECTED_DATA_SHEET_SUFFIX = "_Expected";

    private final AdvancedExcelReader defaultSheetReader = AdvancedExcelReader.builder().build();


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
    public void testMissingRowsAndColumns(String sheetNamePrefix) throws IOException {
        String testDataSheetName = sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX;
        String expectedDataSheetName = sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX;

        SheetConfig config = AdvancedExcelReader.builder()
                .skipHiddenCells(true)
                .buildConfig();

        AdvancedExcelReader sheetReader = new AdvancedExcelReader(config);

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

    private String[][] readMatrix(AdvancedExcelReader sheetReader, ExcelReadRequest request) throws IOException {
        return sheetReader.readSheet(request).getMatrix();
    }

    private String[][] readDefaultMatrix(ExcelReadRequest request) throws IOException {
        return readMatrix(defaultSheetReader, request);
    }
}
