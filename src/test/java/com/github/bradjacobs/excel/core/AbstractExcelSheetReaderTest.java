/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.api.ExcelSheetReader;
import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.config.SanitizeType;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import com.github.bradjacobs.excel.util.TestExcelFileSheetUtils;
import com.github.bradjacobs.excel.util.TestResourceUtil;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.function.Executable;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.FieldSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.UnknownHostException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.config.SanitizeType.SPACES;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

// TODO - this test needs more love
//.  and should eventually replace the 'ExcelSheetReaderTest'
@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public abstract class AbstractExcelSheetReaderTest<T extends ExcelSheetReader, B extends AbstractExcelSheetReader.AbstractSheetConfigBuilder<T, B>> {
    private static final String TEST_DATA_FILE = "testSheetData.xlsx";
    private static final Path TEST_FILE = TestResourceUtil.getResourceFilePath(TEST_DATA_FILE);


    private final T defaultSheetReader = createBuilder().build();

    abstract protected B createBuilder();

    private ExcelSheetReadRequest createRequest(Path path) {
        return ExcelSheetReadRequest.from(path).build();
    }
    private ExcelSheetReadRequest createRequest(File filePath) {
        return ExcelSheetReadRequest.from(filePath).build();
    }
    private ExcelSheetReadRequest createRequest(Path path, String sheetName) {
        return ExcelSheetReadRequest.from(path).byNames(sheetName).build();
    }
    private ExcelSheetReadRequest createRequest(Path path, int sheetIndex) {
        return ExcelSheetReadRequest.from(path).byIndexes(sheetIndex).build();
    }


    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class GeneralTests {


        @Test
        public void basicRead() throws IOException {
            Path testFile = TestResourceUtil.getResourceFilePath("test_data.xlsx");
            ExcelSheetReadRequest req = ExcelSheetReadRequest
                    .from(testFile)
                    .build();

            SheetContent sheetContent = defaultSheetReader.readSheet(req);
            String[][] returnedSheetValues = sheetContent.getMatrix();
            assertArrayEquals(EXPECTED_TEST_DATA, returnedSheetValues, "Mismatch expected sheet content");
        }

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

            ExcelSheetReadRequest req = ExcelSheetReadRequest
                    .from(TEST_FILE)
                    .byNames("GrowingColumnLength")
                    .build();

            List<SheetContent> sheetContentList = defaultSheetReader.readSheets(req);
            SheetContent sheetContent = sheetContentList.get(0);
            String[][] dataMatrix = sheetContent.getMatrix();

            for (String[] rowValues : dataMatrix) {
                assertEquals(3, rowValues.length);
            }
        }

        @Test
        public void readBlankSheet() throws IOException {
            ExcelSheetReadRequest req = createRequest(TEST_FILE, "BlankSheet");
            String[][] dataMatrix = defaultSheetReader.readSheet(req).getMatrix();
            assertEquals(0, dataMatrix.length, "Mismatch expected row count");
        }

        /**
         * Read Sheet that contains hidden blank rows where
         *   row.getLastCellNum() == -1
         * Ensure sheet is still read correctly when these rows are encountered.
         */
        @Test
        public void negativeRowCellNumberRepro() throws IOException {
            ExcelSheetReadRequest req = createRequest(TEST_FILE, "BadRow");
            String[][] csvMatrix = defaultSheetReader.readSheet(req).getMatrix();
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
            ExcelSheetReadRequest req = createRequest(TEST_FILE, "LastColWhitespace");
            String[][] dataMatrix = defaultSheetReader.readSheet(req).getMatrix();
            assertEquals(1, dataMatrix[0].length, "mismatch of expected number of columns in csv output.");
        }


        @Test
        public void readMultipleSheets() throws IOException {

            List<String> sheetNames = List.of("GrowingColumnLength", "withTwoBlankColumns", "WithThreeBlankRows");
            List<String> expectedFirstCell = List.of("aa", "aa11", "Name");

            ExcelSheetReadRequest req = ExcelSheetReadRequest.from(TEST_FILE).byNames(sheetNames).build();

            List<SheetContent> sheetContentList = defaultSheetReader.readSheets(req);
            assertEquals(sheetNames.size(), sheetContentList.size(), "Expected 3 sheets");

            for (int i = 0; i < sheetContentList.size(); i++) {
                assertEquals(expectedFirstCell.get(i), sheetContentList.get(i).getMatrix()[0][0]);
            }
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class AutoTrimTests {
        @Test
        public void autoTrimEnabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            ExcelSheetReadRequest req = createRequest(testFile);

            T sheetReader = createBuilder().autoTrim(true).build();
            SheetContent sheetContent = sheetReader.readSheet(req);
            String[][] dataMatrix = sheetContent.getMatrix();
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDisabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            ExcelSheetReadRequest req = createRequest(testFile);

            T sheetReader = createBuilder().autoTrim(false).build();
            SheetContent sheetContent = sheetReader.readSheet(req);
            String[][] dataMatrix = sheetContent.getMatrix();
            assertEquals(testValue, dataMatrix[0][0]);
        }

        @Test
        public void autoTrimDefault(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            ExcelSheetReadRequest req = createRequest(testFile);

            SheetContent sheetContent = defaultSheetReader.readSheet(req);
            String[][] dataMatrix = sheetContent.getMatrix();
            assertEquals(testValue.trim(), dataMatrix[0][0]);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankRowTests {
        /**
         * Compare the diff of row counts with
         *  skipBlankRows = true vs. false
         */
        @Test
        public void skipBlankRows() throws IOException {
            T sheetReader1 = createBuilder().skipBlankRows(false).build();
            T sheetReader2 = createBuilder().skipBlankRows(true).build();

            ExcelSheetReadRequest req = createRequest(TEST_FILE, "WithThreeBlankRows");

            String[][] dataMatrixRetainBlankRows = sheetReader1.readSheet(req).getMatrix();
            String[][] dataMatrixSkipBlankRows = sheetReader2.readSheet(req).getMatrix();
            int rowDifference = dataMatrixRetainBlankRows.length - dataMatrixSkipBlankRows.length;
            assertEquals(3, rowDifference, "Mismatch expected row count");
        }

        @Test
        public void defaultRetainBlankRows() throws IOException {
            // by default, we keep the blank rows.
            B builder = createBuilder();
            T sheetReader = builder.skipBlankRows(false).build();

            ExcelSheetReadRequest req = createRequest(TEST_FILE, "WithThreeBlankRows");

            String[][] dataMatrixRetainBlankRows = sheetReader.readSheet(req).getMatrix();
            String[][] dataMatrixDefault = defaultSheetReader.readSheet(req).getMatrix();
            assertEquals(dataMatrixRetainBlankRows.length, dataMatrixDefault.length, "Mismatch expected row count");
        }

        /**
         * Regardless of config settings always remove trailing blank rows
         *   (i.e. the last row should contain some values for a non-blank sheet)
         */
        @Test
        public void pruneExtraBlankRows() throws IOException {
            B builder = createBuilder();
            T sheetReader = builder.skipBlankRows(true).build();

            ExcelSheetReadRequest req = createRequest(TEST_FILE, "ExtraBlankRowsAfterData");

            String[][] dataMatrix = sheetReader.readSheet(req).getMatrix();
            assertEquals(2, dataMatrix.length, "Mismatch expected row count");
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankColumnTests {
        /**
         * Compare the diff of column counts with
         *  skipBlankColumns = true vs. false
         */
        @Test
        public void skipBlankColumns() throws IOException {
            T sheetReader1 = createBuilder().skipBlankColumns(false).build();
            T sheetReader2 = createBuilder().skipBlankColumns(true).build();

            ExcelSheetReadRequest req = createRequest(TEST_FILE, "WithTwoBlankColumns");

            String[][] dataMatrixRetainBlankColumns = sheetReader1.readSheet(req).getMatrix();
            String[][] dataMatrixSkipBlankColumns = sheetReader2.readSheet(req).getMatrix();
            int columnDifference = dataMatrixRetainBlankColumns[0].length - dataMatrixSkipBlankColumns[0].length;
            assertEquals(2, columnDifference, "Mismatch expected column count");
        }

        @Test
        public void defaultRetainBlankColumns() throws IOException {
            // by default, we keep the blank columns.
            B builder = createBuilder();
            T sheetReader = builder.skipBlankColumns(false).build();

            ExcelSheetReadRequest req = createRequest(TEST_FILE, "WithTwoBlankColumns");

            String[][] dataMatrixRetainBlankColumns = sheetReader.readSheet(req).getMatrix();
            String[][] dataMatrixDefault = defaultSheetReader.readSheet(req).getMatrix();
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
         * data should look like if it were to skip/ignore the 'hidden data'
         * from the TestDataSheet
         * @param sheetNamePrefix sheetPrefix
         */
        @ParameterizedTest(name = "HiddenTest {index}: Sheet = {0}")
        @ValueSource(strings = {
                "LastInvisible",
                "MiddleBlankColumns",
                "MiddleBlankRows",
                "SingleHIddenRow",
                "SingleHIddenColumn",
                "BaseCase",
                "LastColumn",
                "FirstLastRow",
                "AllColumnHidden",
                "FirstLastColumn",
                "Multi",
                "LastValueInvisible",
                "LongestRowInvisible"
        })
        public void testMissingRowsAndColumns(String sheetNamePrefix) throws IOException {
            String testDataSheetName = sheetNamePrefix + INPUT_DATA_SHEET_SUFFIX;
            String expectedDataSheetName = sheetNamePrefix + EXPECTED_DATA_SHEET_SUFFIX;

            B builder = createBuilder();
            T sheetReader = builder.skipInvisibleCells(true).build();

            ExcelSheetReadRequest req = createRequest(HIDDEN_CELLS_FILE, testDataSheetName);
            ExcelSheetReadRequest expectedReq = createRequest(HIDDEN_CELLS_FILE, expectedDataSheetName);

            String[][] actualMatrix = sheetReader.readSheet(req).getMatrix();
            if (actualMatrix.length > 0) {
                String[] firstRow = actualMatrix[0];
                // check to make sure we don't have rows that contain zero-length arrays (need only check the first)
                assertTrue(firstRow.length > 0, "Matrix return rows with zero-length arrays");
            }

            String[][] expectedMatrix = defaultSheetReader.readSheet(expectedReq).getMatrix();
            assertArrayEquals(expectedMatrix, actualMatrix);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class SanitizeTests {
        private final List<Arguments> sanitizeCases = Arrays.asList(
                // sanitizeType, originalValue, sanitizedValue, isDefaultEnabled
                arguments(named("nbsp-spaces", SPACES), "aa_\u00a0_bb", "aa_ _bb", true),
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

            ExcelSheetReadRequest req = createRequest(testFile);

            // ensure that each reader returns correct expected value.
            String[][] enabledMatrix = enabledSheetReader.readSheet(req).getMatrix();
            String[][] disabledMatrix = disabledSheetReader.readSheet(req).getMatrix();
            String[][] defaultMatrix = defaultSheetReader.readSheet(req).getMatrix();
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

    /**
     *   By default, the DataFormatter uses 'false' for 1904DateWindowing.
     * And this is often correct for most files.  However, if encounter
     * an Excel workbook where 1904DateWindowing is true, then the
     * default behavior will cause dates to be formatted incorrectly.
     *   This test is to verify bug fix where the date values
     * would be incorrect for the 'AdvancedExcelSheetReader' class.
     */
    @Test
    public void test1904Windowing() throws IOException {
        Path date1904File = TestResourceUtil.getResourceFilePath("date1904.xlsx");

        T sheetReader = createBuilder().build();

        ExcelSheetReadRequest req = createRequest(date1904File);
        String[][] dataMatrix = sheetReader.readSheet(req).getMatrix();

        // check the values in column B
        assertEquals("Due Date", dataMatrix[0][1]);
        assertEquals("TBD", dataMatrix[1][1]);
        assertEquals("3/15/08", dataMatrix[2][1]);
        assertEquals("4/30/08", dataMatrix[3][1]);
    }


    private static final String TEXT_FILE_PATH = getResourceFilePath("fake.txt").toAbsolutePath().toString();
    private static final String DIR_PATH = Paths.get("").toAbsolutePath().toString();

    private static final String INPUT_FILE_NOT_FOUND_PATH = "/bogus/path/here/file.xlsx";
    private static final String INPUT_FILE_NOT_FOUND_URL = "https://www.zxfake12.com/foo/bar.html";


    protected static List<Arguments> invalidInputPaths() {
        return Arrays.asList(
                arguments(named("File Not Found", INPUT_FILE_NOT_FOUND_PATH), FileNotFoundException.class, "Invalid Excel file path: /bogus/path/here/file.xlsx"),
                arguments(named("Null Input", null), IllegalArgumentException.class, "Either file path or url must be provided"),
                arguments(named("Invalid Excel File Input", TEXT_FILE_PATH), IOException.class, null),
                arguments(named("Directory Input", DIR_PATH), IllegalArgumentException.class, "The input file cannot be a directory.")
        );
    }

    protected static List<Arguments> invalidInputUrls() {
        return Arrays.asList(
                arguments(named("Url Not Found", INPUT_FILE_NOT_FOUND_URL), UnknownHostException.class, null),
                arguments(named("Null Input", null), IllegalArgumentException.class, "Either file path or url must be provided"),
                arguments(named("Directory Input", "file:///"), IllegalArgumentException.class, "The input file cannot be a directory."),
                arguments(named("Invalid Url Protocol", "jar:file:/C:/foo/jar/parser.jar!/test.xlsx"), IllegalArgumentException.class, "URL has an unsupported protocol: jar")
        );
    }


    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class ErrorHandlingTests {
        @Test
        public void nullInputStream() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                sheetReader.readSheet(null);
            });
        }

        @Test
        public void negativeSheetIndex() {
            // todo - check message
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                ExcelSheetReadRequest req = createRequest(Path.of("fake.xlsx"), -2);
                sheetReader.readSheet(req);
            });
        }

        @Test
        public void nullSheetName() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                // todo check message
                T sheetReader = createBuilder().build();
                // fake inputStream, but expect the negative index to be detected first.
                ExcelSheetReadRequest req = createRequest(Path.of("fake.xlsx"), null);
                sheetReader.readSheet(req);
            });
        }

        @Test
        public void emptySheetName() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                // todo check message
                T sheetReader = createBuilder().build();
                // fake inputStream, but expect the negative index to be detected first.
                ExcelSheetReadRequest req = createRequest(TEST_FILE, "");
                sheetReader.readSheet(req);
            });
        }

        @Test
        public void unknownSheetName() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                // todo check message
                T sheetReader = createBuilder().build();
                ExcelSheetReadRequest req = createRequest(TEST_FILE, "UnknownSheetName");
                sheetReader.readSheet(req);
            });

            assertEquals("Requested Excel sheet not found: 'UnknownSheetName'", thrown.getMessage());
        }
    }

// if (FileMagic.OLE2 != fm) {
//     throw new IOException("Can't open workbook - unsupported file type: "+fm);
// }

    @ParameterizedTest
    @MethodSource("invalidInputPaths")
    public void invalidInputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws IOException {
        Path path = location != null ? Paths.get(location) : null;
        T sheetReader = createBuilder().build();

        Executable executable = () -> {
            ExcelSheetReadRequest req = createRequest(path);
            sheetReader.readSheet(req);
        };
        assertExecutableException(executable, expectedException, expectedMessage);
    }

    @ParameterizedTest
    @MethodSource("invalidInputUrls")
    public void invalidInputUrlParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws IOException {
        URL url = location != null ? new URL(location) : null;
        T sheetReader = createBuilder().build();

        Executable executable = () -> {
            ExcelSheetReadRequest req = ExcelSheetReadRequest.from(url).build();
            sheetReader.readSheet(req);
        };
        assertExecutableException(executable, expectedException, expectedMessage);
    }



    /**
     * Junit run the executable and check that it throws the expected exception and msg.
     * @param executable executable
     * @param expectedException expectedException
     * @param expectedMessage expectedMessage
     */
    private void assertExecutableException(
            Executable executable,
            Class<? extends Exception> expectedException,
            String expectedMessage) {
        Exception exception = assertThrows(Exception.class, executable);
        assertEquals(expectedException, exception.getClass(), "Mismatch expected exception thrown");
        if (expectedMessage != null) {
            assertEquals(expectedMessage, exception.getMessage());
        }
    }


    private static final String[][] EXPECTED_TEST_DATA = new String[][] {
            {"col1","col_2","COL3","Col4","Col5"},
            {"Text","Space Text","$_() ' extra chars","92\"82_cell_w_quote","extra"},
            {"11","22","33","66","row_contains_formulas"},
            {"with_dates_row","9/8/20","9/9/20","1999-07-04",""},
            {"sparse_column_row","","","val",""},
            {"","11:33","11:33","",""},
            {"","","","",""},
            {"","","","",""},
            {"misc_string-row","value     with     spaces","funky\"quote","value\twith\ttabs",""},
            {"misc_number_row","-1","4.765400","1.23",""},
            {"invalid_formula_row","#VALUE!","#VALUE!","#NAME?","i"},
            {"numeric_row","1.888888889","3.14","-0.57","1.1"}
    };
}
