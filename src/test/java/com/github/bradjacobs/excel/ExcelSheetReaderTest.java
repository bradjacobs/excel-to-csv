/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;
import com.github.bradjacobs.excel.util.TestResourceUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.FieldSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.DASHES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.SPACES;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

public class ExcelSheetReaderTest {

    private static final String TEST_DATA_FILE = "testSheetData.xlsx";

    private static final ExcelSheetReader DEFAULT_SHEET_READER =
            ExcelSheetReader.builder().build();

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class GeneralTests {
        /**
         * Ensure that the final matrix column counts
         * are the same for each row.  The last column
         * should have a value for at least 1 of the rows
         */
        @Test
        public void ensureColumnSize() {
            // Test Data is:
            //     aa
            //     bb,bb
            //     ccc,ccc,ccc
            // ALL rows should have a length of 3
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "GrowingColumnLength");
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            for (String[] rowValues : dataMatrix) {
                assertEquals(3, rowValues.length);
            }
        }

        @Test
        public void readBlankSheet() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "BlankSheet");
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals(0, dataMatrix.length, "Mismatch expected row count");
        }

        /**
         * Read Sheet that contains hidden blank rows where
         *   row.getLastCellNum() == -1
         * Ensure sheet is still read correctly when these rows are encountered.
         */
        @Test
        public void negativeRowCellNumberRepro() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "BadRow");
            String[][] csvMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals("aaa", csvMatrix[0][0]);
            assertEquals("bbb", csvMatrix[0][1]);
            assertEquals("ccc", csvMatrix[2][0]);
            assertEquals("ddd", csvMatrix[2][1]);
        }

        /**
         * Sheet where the last column is filled with different
         * whitespace characters.  Since we have default values of:
         *   sanitizeSpaces=true AND autoTrimSpaces=true
         * Then this last column should be removed from the result.
         */
        @Test
        public void handleExtraBlankWhitespaceColumn() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "LastColWhitespace");
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals(1, dataMatrix[0].length, "mismatch of expected number of columns in csv output.");
        }

        /**
         * Null Sheet parameter should be IllegalArgumentException
         */
        @Test
        public void nullSheetParamCheck() {
            assertThrows(IllegalArgumentException.class, () -> {
                DEFAULT_SHEET_READER.convertToDataMatrix(null);
            });
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class AutoTrimTests {
        @Test
        public void autoTrimEnabled(@TempDir Path tempDir) {
            String testValue = "  aa bb  ";
            Sheet testSheet = createSingleCellExcelSheet(tempDir, testValue);
            ExcelSheetReader excelSheetReader1 = ExcelSheetReader.builder()
                    .autoTrim(true)
                    .build();
            String[][] dataMatrix = excelSheetReader1.convertToDataMatrix(testSheet);
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDisabled(@TempDir Path tempDir) {
            String testValue = "  aa bb  ";
            Sheet testSheet = createSingleCellExcelSheet(tempDir, testValue);
            ExcelSheetReader excelSheetReader1 = ExcelSheetReader.builder()
                    .autoTrim(false)
                    .build();
            String[][] dataMatrix = excelSheetReader1.convertToDataMatrix(testSheet);
            assertEquals(testValue, dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDefault(@TempDir Path tempDir) {
            String testValue = "  aa bb  ";
            Sheet testSheet = createSingleCellExcelSheet(tempDir, testValue);
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankRowTests {
        /**
         * Compare the diff of row counts with
         *  removeBlankRows = true vs false
         */
        @Test
        public void removeBlankRows() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithThreeBlankRows");
            ExcelSheetReader excelSheetReader1 = ExcelSheetReader.builder()
                    .removeBlankRows(false)
                    .build();
            ExcelSheetReader excelSheetReader2 = ExcelSheetReader.builder()
                    .removeBlankRows(true)
                    .build();
            String[][] dataMatrixRetainBlankRows = excelSheetReader1.convertToDataMatrix(testSheet);
            String[][] dataMatrixRemoveBlankRows = excelSheetReader2.convertToDataMatrix(testSheet);
            int rowDifference = dataMatrixRetainBlankRows.length - dataMatrixRemoveBlankRows.length;
            assertEquals(3, rowDifference, "Mismatch expected row count");
        }

        @Test
        public void defaultRetainBlankRows() {
            // by default we keep the blank rows.
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithThreeBlankRows");
            ExcelSheetReader excelSheetReader1 = ExcelSheetReader.builder()
                    .removeBlankRows(false)
                    .build();
            String[][] dataMatrixRetainBlankRows = excelSheetReader1.convertToDataMatrix(testSheet);
            String[][] dataMatrixDefault = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals(dataMatrixRetainBlankRows.length, dataMatrixDefault.length, "Mismatch expected row count");
        }

        /**
         * Regardless of setting always remove trailing blank rows
         *   (i.e. the last row should contain some values for a non-blank sheet)
         */
        @Test
        public void pruneExtraBlankRows() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "ExtraBlankRowsAfterData");
            ExcelSheetReader excelSheetReader = ExcelSheetReader.builder()
                    .removeBlankRows(false)
                    .build();
            String[][] dataMatrix = excelSheetReader.convertToDataMatrix(testSheet);
            assertEquals(dataMatrix.length, 2, "Mismatch expected row count");
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankColumnTests {
        /**
         * Compare the diff of column counts with
         *  removeBlankColumns = true vs false
         */
        @Test
        public void removeBlankColumns() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithTwoBlankColumns");
            ExcelSheetReader excelSheetReader1 = ExcelSheetReader.builder()
                    .removeBlankColumns(false)
                    .build();
            ExcelSheetReader excelSheetReader2 = ExcelSheetReader.builder()
                    .removeBlankColumns(true)
                    .build();
            String[][] dataMatrixRetainBlankColumns = excelSheetReader1.convertToDataMatrix(testSheet);
            String[][] dataMatrixRemoveBlankColumns = excelSheetReader2.convertToDataMatrix(testSheet);
            int columnDifference = dataMatrixRetainBlankColumns[0].length - dataMatrixRemoveBlankColumns[0].length;
            assertEquals(2, columnDifference, "Mismatch expected column count");
        }

        @Test
        public void defaultRetainBlankCoumns() {
            // by default we keep the blank columns.
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithTwoBlankColumns");
            ExcelSheetReader excelSheetReader1 = ExcelSheetReader.builder()
                    .removeBlankColumns(false)
                    .build();
            String[][] dataMatrixRetainBlankColumns = excelSheetReader1.convertToDataMatrix(testSheet);
            String[][] dataMatrixDefault = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals(dataMatrixRetainBlankColumns[0].length, dataMatrixDefault[0].length, "Mismatch expected column count");
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class InvisibleCellsTests {
        private static final String HIDDEN_CELLS_DATA_FILE = "skipHiddenTestData.xlsx";
        private static final String INPUT_DATA_SHEET_SUFFIX = "_Data";
        private static final String EXPECTED_DATA_SHEET_SUFFIX = "_Expected";

        /**
         * The HiddenCellsDataFile contains multiple sheets in set of 2.
         *     (TestDataSheet, ExpectedResultsDataSheet)
         * where the ExpectedResultsDataSheet represents what the
         * data should look like if were to remove/ignore the 'hidden data'
         * from the TestDataSheet
         * @param sheetNamePrefix sheetPrefix
         */
        @ParameterizedTest(name = "HiddenTest {index}: Sheet = {0}")
        @ValueSource(strings = {
                "BaseCase",
                "LastColumn",
                "FirstLastRow",
                "AllColumnHidden",
                "FirstLastColumn",
                "Multi"})
        public void testMissingRowsAndColumns(String sheetNamePrefix) {
            String testDataSheetName = sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX;
            String expectedDataSheetName = sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX;

            Sheet testDataSheet = getFileSheet(HIDDEN_CELLS_DATA_FILE, testDataSheetName);
            Sheet expectedDataSheet = getFileSheet(HIDDEN_CELLS_DATA_FILE, expectedDataSheetName);

            ExcelSheetReader removeHiddenSheetReader = ExcelSheetReader.builder()
                    .removeInvisibleCells(true)
                    .build();

            String[][] actualMatrix = removeHiddenSheetReader.convertToDataMatrix(testDataSheet);
            if (actualMatrix.length > 0) {
                String[] firstRow = actualMatrix[0];
                // check to make sure we don't have rows that contain zero-length arrays (need only check the first)
                assertTrue(firstRow.length > 0, "Matrix return rows with zero-length arrays");
            }

            String[][] expectedMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(expectedDataSheet);
            assertArrayEquals(expectedMatrix, actualMatrix);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class SanitizeTests {
        private final List<Arguments> sanitizeCases = Arrays.asList(
                // flagType, originalValue, sanitizedValue, isDefaultEnabled
                arguments(named("nbsp-spaces", SPACES), "aa\u00a0bb", "aa bb", true),
                arguments(named("doubleSmart-quotes", QUOTES), "with_“x”_val", "with_\"x\"_val", true),
                arguments(named("singleSmart-quotes", QUOTES), "‘hi’", "'hi'", true),
                arguments(named("emDash-dashes", DASHES), "aa—bb", "aa-bb", false),
                arguments(named("facade-diacritics", BASIC_DIACRITICS),  "Façade", "Facade", false)
        );

        @ParameterizedTest
        @FieldSource("sanitizeCases")
        public void sanitizeSheet(
                CharSanitizeFlag flag,
                String origValue,
                String sanitizedValue,
                boolean isDefaultEnabled,
                @TempDir Path tempDir) {
            // generate an Excel File Sheet with 1 single cell value.
            Sheet testSheet = createSingleCellExcelSheet(tempDir, origValue);

            // create readers set to both enabled and disabled
            ExcelSheetReader enabledSheetReader = createSanitizeSheetReader(flag, true);
            ExcelSheetReader disabledSheetReader = createSanitizeSheetReader(flag, false);

            // ensure that each reader returns correct expected value.
            String[][] enabledMatrix = enabledSheetReader.convertToDataMatrix(testSheet);
            String[][] disabledMatrix = disabledSheetReader.convertToDataMatrix(testSheet);
            String[][] defaultMatrix = DEFAULT_SHEET_READER.convertToDataMatrix(testSheet);
            assertEquals(sanitizedValue, enabledMatrix[0][0]);
            assertEquals(origValue, disabledMatrix[0][0]);

            if (isDefaultEnabled) {
                assertEquals(enabledMatrix[0][0], defaultMatrix[0][0]);
            }
            else {
                assertEquals(disabledMatrix[0][0], defaultMatrix[0][0]);
            }
        }

        private ExcelSheetReader createSanitizeSheetReader(CharSanitizeFlag flag, boolean enabled) {
            ExcelSheetReader.Builder builder = ExcelSheetReader.builder();
            switch (flag) {
                case SPACES:
                    return builder.sanitizeSpaces(enabled).build();
                case QUOTES:
                    return builder.sanitizeQuotes(enabled).build();
                case DASHES:
                    return builder.sanitizeDashes(enabled).build();
                case BASIC_DIACRITICS:
                    return builder.sanitizeDiacritics(enabled).build();
                default:
                    throw new IllegalArgumentException("Unabled CharSanitizeFlag: " + flag);
            }
        }
    }

    private Sheet getFileSheet(String fileName, String sheetName) {
        InputStream inputStream = TestResourceUtil.readResourceFileInputStream(fileName);
        try (inputStream; Workbook wb = WorkbookFactory.create(inputStream)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet was not found: " + sheetName);
            }
            return sheet;
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to read Sheet from Excel File: " + e.getMessage(), e);
        }
    }

    /**
     * Create a test Excell file that contains 1 sheet with 1 cell value.
     *   NOTE: assume that test @TempDir will automatically handle file clean up.
     * @param tempDir temp directory where the excel file should live
     * @param cellValue value to set for cell
     * @return Excel Sheet
     */
    private Sheet createSingleCellExcelSheet(Path tempDir, String cellValue) {
        Path testFile = tempDir.resolve("tmp_"+System.currentTimeMillis()+".xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("TestCellSheet");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(cellValue);
        try (FileOutputStream out = new FileOutputStream(testFile.toFile())) {
            workbook.write(out);
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to create Temp Excel File: " + e.getMessage(), e);
        }

        try (InputStream inputStream = Files.newInputStream(testFile);
             Workbook wb = WorkbookFactory.create(inputStream)) {
            return wb.getSheetAt(0);
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to read Sheet from Excel File: " + e.getMessage(), e);
        }
    }
}
