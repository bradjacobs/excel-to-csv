/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.config.SanitizeType;
import com.github.bradjacobs.excel.util.TestExcelFileSheetUtils;
import com.github.bradjacobs.excel.util.TestResourceUtil;
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
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

// TODO - this test needs more love
//.  and should evenually replace the 'ExcelSheetReaderTest'
abstract public class AbstractExcelSheetReaderTest<T extends ExcelSheetReader, B extends AbstractExcelSheetReader.AbstractSheetConfigBuilder<T, B>> {
    private static final String TEST_DATA_FILE = "testSheetData.xlsx";
    private static final Path TEST_FILE = TestResourceUtil.getResourceFilePath(TEST_DATA_FILE);

    private final T defaultSheetReader = createBuilder().build();

    abstract protected B createBuilder();

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class GeneralTests {
        /**
         * Ensure that the final matrix column counts
         * are the same for each row.  The last column
         * should have a value for at least 1 of the rows
         */
        @Test
        public void ensureColumnSize() throws IOException {
            // Test Data is:
            //     aa
            //     bb,bb
            //     ccc,ccc,ccc
            // ALL rows should have a length of 3
            String[][] dataMatrix = defaultSheetReader.readExcelSheetData(TEST_FILE, "GrowingColumnLength");
            for (String[] rowValues : dataMatrix) {
                assertEquals(3, rowValues.length);
            }
        }

        @Test
        public void readBlankSheet() throws IOException {
            String[][] dataMatrix = defaultSheetReader.readExcelSheetData(TEST_FILE, "BlankSheet");
            assertEquals(0, dataMatrix.length, "Mismatch expected row count");
        }

        /**
         * Read Sheet that contains hidden blank rows where
         *   row.getLastCellNum() == -1
         * Ensure sheet is still read correctly when these rows are encountered.
         */
        @Test
        public void negativeRowCellNumberRepro() throws IOException {
            String[][] csvMatrix = defaultSheetReader.readExcelSheetData(TEST_FILE, "BadRow");
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
        public void handleExtraBlankWhitespaceColumn() throws IOException {
            String[][] dataMatrix = defaultSheetReader.readExcelSheetData(TEST_FILE, "LastColWhitespace");
            assertEquals(1, dataMatrix[0].length, "mismatch of expected number of columns in csv output.");
        }

        // todo - fix below

//        /**
//         * Null Sheet parameter should be IllegalArgumentException
//         */
//        @Test
//        public void nullSheetParamCheck() {
//            assertThrows(IllegalArgumentException.class, () -> {
//                defaultSheetDataExtractor.readExcelSheetData(null);
//            });
//        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class AutoTrimTests {
        @Test
        public void autoTrimEnabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);

            T sheetReader1 = createBuilder().autoTrim(true).build();
            String[][] dataMatrix = sheetReader1.readExcelSheetData(testFile, 0);
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDisabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);

            T sheetReader1 = createBuilder().autoTrim(false).build();
            String[][] dataMatrix = sheetReader1.readExcelSheetData(testFile, 0);
            assertEquals(testValue, dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDefault(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            String[][] dataMatrix = defaultSheetReader.readExcelSheetData(testFile, 0);
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankRowTests {
        /**
         * Compare the diff of row counts with
         *  removeBlankRows = true vs. false
         */
        @Test
        public void removeBlankRows() throws IOException {
            T sheetReader1 = createBuilder().removeBlankRows(false).build();
            T sheetReader2 = createBuilder().removeBlankRows(true).build();

            String[][] dataMatrixRetainBlankRows = sheetReader1.readExcelSheetData(TEST_FILE, "WithThreeBlankRows");
            String[][] dataMatrixRemoveBlankRows = sheetReader2.readExcelSheetData(TEST_FILE, "WithThreeBlankRows");
            int rowDifference = dataMatrixRetainBlankRows.length - dataMatrixRemoveBlankRows.length;
            assertEquals(3, rowDifference, "Mismatch expected row count");
        }

        @Test
        public void defaultRetainBlankRows() throws IOException {
            // by default we keep the blank rows.
            B builder = createBuilder();
            T sheetReader = builder.removeBlankRows(false).build();

            String[][] dataMatrixRetainBlankRows = sheetReader.readExcelSheetData(TEST_FILE, "WithThreeBlankRows");
            String[][] dataMatrixDefault = defaultSheetReader.readExcelSheetData(TEST_FILE, "WithThreeBlankRows");
            assertEquals(dataMatrixRetainBlankRows.length, dataMatrixDefault.length, "Mismatch expected row count");
        }

        /**
         * Regardless of setting always remove trailing blank rows
         *   (i.e. the last row should contain some values for a non-blank sheet)
         */
        @Test
        public void pruneExtraBlankRows() throws IOException {
            B builder = createBuilder();
            T sheetReader = builder.removeBlankRows(true).build();

            String[][] dataMatrix = sheetReader.readExcelSheetData(TEST_FILE, "ExtraBlankRowsAfterData");
            assertEquals(2, dataMatrix.length, "Mismatch expected row count");
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankColumnTests {
        /**
         * Compare the diff of column counts with
         *  removeBlankColumns = true vs. false
         */
        @Test
        public void removeBlankColumns() throws IOException {
            T sheetReader1 = createBuilder().removeBlankColumns(false).build();
            T sheetReader2 = createBuilder().removeBlankColumns(true).build();

            String[][] dataMatrixRetainBlankColumns = sheetReader1.readExcelSheetData(TEST_FILE, "WithTwoBlankColumns");
            String[][] dataMatrixRemoveBlankColumns = sheetReader2.readExcelSheetData(TEST_FILE, "WithTwoBlankColumns");
            int columnDifference = dataMatrixRetainBlankColumns[0].length - dataMatrixRemoveBlankColumns[0].length;
            assertEquals(2, columnDifference, "Mismatch expected column count");
        }

        @Test
        public void defaultRetainBlankColumns() throws IOException {
            // by default we keep the blank columns.
            B builder = createBuilder();
            T sheetReader = builder.removeBlankColumns(false).build();

            String[][] dataMatrixRetainBlankColumns = sheetReader.readExcelSheetData(TEST_FILE, "WithTwoBlankColumns");
            String[][] dataMatrixDefault = defaultSheetReader.readExcelSheetData(TEST_FILE, "WithTwoBlankColumns");
            assertEquals(dataMatrixRetainBlankColumns[0].length, dataMatrixDefault[0].length, "Mismatch expected column count");
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class InvisibleCellsTests {
        private static final String HIDDEN_CELLS_DATA_FILE = "skipHiddenTestData.xlsx";
        private final Path HIDDEN_CELLS_FILE = TestResourceUtil.getResourceFilePath(HIDDEN_CELLS_DATA_FILE);
        private static final String INPUT_DATA_SHEET_SUFFIX = "_Data";
        private static final String EXPECTED_DATA_SHEET_SUFFIX = "_Expected";

        /**
         * The HiddenCellsDataFile contains multiple sheets in set of 2.
         *     (TestDataSheet, ExpectedResultsDataSheet)
         * where the ExpectedResultsDataSheet represents what the
         * data should look like if it were to remove/ignore the 'hidden data'
         * from the TestDataSheet
         * @param sheetNamePrefix sheetPrefix
         */
        @ParameterizedTest(name = "HiddenTest {index}: Sheet = {0}")
        @ValueSource(strings = {
                "LastInvisible",
                "MiddleBlankColumns",
                "MiddleBlankRows",
                "BaseCase",
                "LastColumn",
                "FirstLastRow",
                "AllColumnHidden",
                "FirstLastColumn",
                "Multi",
                "LastValueInvisible",
                "LongestRowInvisible"})
        public void testMissingRowsAndColumns(String sheetNamePrefix) throws IOException {
            String testDataSheetName = sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX;
            String expectedDataSheetName = sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX;

            B builder = createBuilder();
            T sheetReader = builder.removeInvisibleCells(true).build();

            String[][] actualMatrix = sheetReader.readExcelSheetData(HIDDEN_CELLS_FILE, testDataSheetName);
            if (actualMatrix.length > 0) {
                String[] firstRow = actualMatrix[0];
                // check to make sure we don't have rows that contain zero-length arrays (need only check the first)
                assertTrue(firstRow.length > 0, "Matrix return rows with zero-length arrays");
            }

            String[][] expectedMatrix = defaultSheetReader.readExcelSheetData(HIDDEN_CELLS_FILE, expectedDataSheetName);
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
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, origValue);

            // create readers set to both enabled and disabled
            T enabledSheetReader = createSanitizeSheetReader(type, true);
            T disabledSheetReader = createSanitizeSheetReader(type, false);

            // ensure that each reader returns correct expected value.
            String[][] enabledMatrix = enabledSheetReader.readExcelSheetData(testFile, 0);
            String[][] disabledMatrix = disabledSheetReader.readExcelSheetData(testFile, 0);
            String[][] defaultMatrix = defaultSheetReader.readExcelSheetData(testFile, 0);
            assertEquals(sanitizedValue, enabledMatrix[0][0]);
            assertEquals(origValue, disabledMatrix[0][0]);

            if (isDefaultEnabled) {
                assertEquals(enabledMatrix[0][0], defaultMatrix[0][0]);
            }
            else {
                assertEquals(disabledMatrix[0][0], defaultMatrix[0][0]);
            }
        }

        private T createSanitizeSheetReader(SanitizeType type, boolean enabled) {
            B builder = createBuilder();
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
                    throw new IllegalArgumentException("Unhandled SanitizeType: " + type);
            }
        }
    }

}

//. "/var/folders/ll/4cnnvnl1781fl9w10yvgrcr40000gn/T/junit-12636384025942901965/tmp_1772514140336.xlsx"