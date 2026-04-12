/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.config.SanitizeType;
import com.github.bradjacobs.excel.util.TestExcelFileSheetUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.FieldSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.config.SanitizeType.SPACES;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

// TODO
//   this class is now basically duplicated by the
//   newer abstract class tests.
//   Need to update (or remove) this test in a future update.
class OriginalExcelSheetReaderTest {

    private static final String TEST_DATA_FILE = "testSheetData.xlsx";

    private static final StandardExcelSheetReader DEFAULT_SHEET_READER =
            StandardExcelSheetReader.builder().build();

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
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
            for (String[] rowValues : dataMatrix) {
                assertEquals(3, rowValues.length);
            }
        }

        @Test
        public void readBlankSheet() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "BlankSheet");
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
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
            String[][] csvMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
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
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
            assertEquals(1, dataMatrix[0].length, "mismatch of expected number of columns in csv output.");
        }

        /**
         * Null Sheet parameter should be IllegalArgumentException
         */
        @Test
        public void nullSheetParamCheck() {
            assertThrows(IllegalArgumentException.class, () -> {
                DEFAULT_SHEET_READER.convertToSheetContentArray(null);
            });
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class AutoTrimTests {
        @Test
        public void autoTrimEnabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Sheet testSheet = createSingleCellExcelSheet(tempDir, testValue);
            StandardExcelSheetReader standardExcelSheetReader1 = StandardExcelSheetReader.builder()
                    .autoTrim(true)
                    .build();
            String[][] dataMatrix = standardExcelSheetReader1.convertToSheetContentArray(testSheet);
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDisabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Sheet testSheet = createSingleCellExcelSheet(tempDir, testValue);
            StandardExcelSheetReader standardExcelSheetReader1 = StandardExcelSheetReader.builder()
                    .autoTrim(false)
                    .build();
            String[][] dataMatrix = standardExcelSheetReader1.convertToSheetContentArray(testSheet);
            assertEquals(testValue, dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDefault(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Sheet testSheet = createSingleCellExcelSheet(tempDir, testValue);
            String[][] dataMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankRowTests {
        /**
         * Compare the diff of row counts with
         *  skipBlankRows = true vs false
         */
        @Test
        public void skipBlankRows() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithThreeBlankRows");
            StandardExcelSheetReader standardExcelSheetReader1 = StandardExcelSheetReader.builder()
                    .skipBlankRows(false)
                    .build();
            StandardExcelSheetReader standardExcelSheetReader2 = StandardExcelSheetReader.builder()
                    .skipBlankRows(true)
                    .build();
            String[][] dataMatrixRetainBlankRows = standardExcelSheetReader1.convertToSheetContentArray(testSheet);
            String[][] dataMatrixSkipBlankRows = standardExcelSheetReader2.convertToSheetContentArray(testSheet);
            int rowDifference = dataMatrixRetainBlankRows.length - dataMatrixSkipBlankRows.length;
            assertEquals(3, rowDifference, "Mismatch expected row count");
        }

        @Test
        public void defaultRetainBlankRows() {
            // by default we keep the blank rows.
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithThreeBlankRows");
            StandardExcelSheetReader standardExcelSheetReader1 = StandardExcelSheetReader.builder()
                    .skipBlankRows(false)
                    .build();
            String[][] dataMatrixRetainBlankRows = standardExcelSheetReader1.convertToSheetContentArray(testSheet);
            String[][] dataMatrixDefault = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
            assertEquals(dataMatrixRetainBlankRows.length, dataMatrixDefault.length, "Mismatch expected row count");
        }

        /**
         * Regardless of setting always remove trailing blank rows
         *   (i.e. the last row should contain some values for a non-blank sheet)
         */
        @Test
        public void pruneExtraBlankRows() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "ExtraBlankRowsAfterData");
            StandardExcelSheetReader standardExcelSheetReader = StandardExcelSheetReader.builder()
                    .skipBlankRows(false)
                    .build();
            String[][] dataMatrix = standardExcelSheetReader.convertToSheetContentArray(testSheet);
            assertEquals(2, dataMatrix.length, "Mismatch expected row count");
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankColumnTests {
        /**
         * Compare the diff of column counts with
         *  skipBlankColumns = true vs false
         */
        @Test
        public void skipBlankColumns() {
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithTwoBlankColumns");
            StandardExcelSheetReader standardExcelSheetReader1 = StandardExcelSheetReader.builder()
                    .skipBlankColumns(false)
                    .build();
            StandardExcelSheetReader standardExcelSheetReader2 = StandardExcelSheetReader.builder()
                    .skipBlankColumns(true)
                    .build();
            String[][] dataMatrixRetainBlankColumns = standardExcelSheetReader1.convertToSheetContentArray(testSheet);
            String[][] dataMatrixSkipBlankColumns = standardExcelSheetReader2.convertToSheetContentArray(testSheet);
            int columnDifference = dataMatrixRetainBlankColumns[0].length - dataMatrixSkipBlankColumns[0].length;
            assertEquals(2, columnDifference, "Mismatch expected column count");
        }

        @Test
        public void defaultRetainBlankCoumns() {
            // by default we keep the blank columns.
            Sheet testSheet = getFileSheet(TEST_DATA_FILE, "WithTwoBlankColumns");
            StandardExcelSheetReader standardExcelSheetReader1 = StandardExcelSheetReader.builder()
                    .skipBlankColumns(false)
                    .build();
            String[][] dataMatrixRetainBlankColumns = standardExcelSheetReader1.convertToSheetContentArray(testSheet);
            String[][] dataMatrixDefault = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
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
         * data should look like if were to skip/ignore the 'hidden data'
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
                "Multi",
                "LastValueInvisible",
                "LongestRowInvisible"})
        public void testMissingRowsAndColumns(String sheetNamePrefix) {
            String testDataSheetName = sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX;
            String expectedDataSheetName = sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX;

            Sheet testDataSheet = getFileSheet(HIDDEN_CELLS_DATA_FILE, testDataSheetName);
            Sheet expectedDataSheet = getFileSheet(HIDDEN_CELLS_DATA_FILE, expectedDataSheetName);

            StandardExcelSheetReader skipHiddenCellsSheetReader = StandardExcelSheetReader.builder()
                    .skipInvisibleCells(true)
                    .build();

            String[][] actualMatrix = skipHiddenCellsSheetReader.convertToSheetContentArray(testDataSheet);
            if (actualMatrix.length > 0) {
                String[] firstRow = actualMatrix[0];
                // check to make sure we don't have rows that contain zero-length arrays (need only check the first)
                assertTrue(firstRow.length > 0, "Matrix return rows with zero-length arrays");
            }

            String[][] expectedMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(expectedDataSheet);
            assertArrayEquals(expectedMatrix, actualMatrix);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class SanitizeTests {
        private final List<Arguments> sanitizeCases = Arrays.asList(
                // sanitizeType, originalValue, sanitizedValue, isDefaultEnabled
                arguments(named("nbsp-spaces", SPACES), "aa\u00a0bb", "aa bb", true),
                arguments(named("doubleSmart-quotes", QUOTES), "with_“x”_val", "with_\"x\"_val", true),
                arguments(named("singleSmart-quotes", QUOTES), "‘hi’", "'hi'", true),
                arguments(named("emDash-dashes", DASHES), "aa—bb", "aa-bb", false),
                arguments(named("facade-diacritics", BASIC_DIACRITICS),  "Façade", "Facade", false)
        );

        @ParameterizedTest
        @FieldSource("sanitizeCases")
        public void sanitizeSheet(
                SanitizeType type,
                String origValue,
                String sanitizedValue,
                boolean isDefaultEnabled,
                @TempDir Path tempDir) throws IOException {
            // generate an Excel File Sheet with 1 single cell value.
            Sheet testSheet = createSingleCellExcelSheet(tempDir, origValue);

            // create readers set to both enabled and disabled
            StandardExcelSheetReader enabledSheetReader = createSanitizeSheetReader(type, true);
            StandardExcelSheetReader disabledSheetReader = createSanitizeSheetReader(type, false);

            // ensure that each reader returns correct expected value.
            String[][] enabledMatrix = enabledSheetReader.convertToSheetContentArray(testSheet);
            String[][] disabledMatrix = disabledSheetReader.convertToSheetContentArray(testSheet);
            String[][] defaultMatrix = DEFAULT_SHEET_READER.convertToSheetContentArray(testSheet);
            assertEquals(sanitizedValue, enabledMatrix[0][0]);
            assertEquals(origValue, disabledMatrix[0][0]);

            if (isDefaultEnabled) {
                assertEquals(enabledMatrix[0][0], defaultMatrix[0][0]);
            }
            else {
                assertEquals(disabledMatrix[0][0], defaultMatrix[0][0]);
            }
        }

        private StandardExcelSheetReader createSanitizeSheetReader(SanitizeType type, boolean enabled) {
            StandardExcelSheetReader.Builder builder = StandardExcelSheetReader.builder();
            switch (type) {
                case SPACES:
                    return builder.sanitizeSpaces(enabled).build();
                case QUOTES:
                    return builder.sanitizeQuotes(enabled).build();
                case DASHES:
                    return builder.sanitizeDashes(enabled).build();
                case BASIC_DIACRITICS:
                    return builder.sanitizeDiacritics(enabled).build();
                default:
                    throw new IllegalArgumentException("Unabled SanitizeType: " + type);
            }
        }
    }

    private Sheet getFileSheet(String fileName, String sheetName) {
        return TestExcelFileSheetUtils.getFileSheet(fileName, sheetName);
    }

    private Sheet createSingleCellExcelSheet(Path tempDir, String cellValue) throws IOException {
        return TestExcelFileSheetUtils.createSingleCellExcelSheet(tempDir, cellValue);
   }
}
