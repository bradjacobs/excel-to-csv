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
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.function.Executable;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.FieldSource;
import org.junit.jupiter.params.provider.MethodSource;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.UnknownHostException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.config.SanitizeType.SPACES;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileUrl;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
public abstract class AbstractExcelSheetReaderTest<T extends ExcelSheetReader, B extends AbstractExcelSheetReader.AbstractSheetConfigBuilder<T, B>> {
    private static final String TEST_DATA_FILE = "testSheetData.xlsx";
    private static final Path TEST_FILE = TestResourceUtil.getResourceFilePath(TEST_DATA_FILE);

    private static final String GROWING_COLUMN_LENGTH_SHEET = "GrowingColumnLength";
    private static final String WITH_THREE_BLANK_ROWS_SHEET = "WithThreeBlankRows";
    private static final String WITH_TWO_BLANK_COLUMNS_SHEET = "WithTwoBlankColumns";

    private static final String EXPECTED_SHEET_CONTENT_MESSAGE = "Mismatch expected sheet content";
    private static final String EXPECTED_ROW_COUNT_MESSAGE = "Mismatch expected row count";
    private static final String EXPECTED_COLUMN_COUNT_MESSAGE = "Mismatch expected column count";

    private final T defaultSheetReader = createBuilder().build();

    protected abstract B createBuilder();

    private ExcelSheetReadRequest createRequest(Path path) {
        return ExcelSheetReadRequest.from(path).build();
    }

    private ExcelSheetReadRequest createRequest(File filePath) {
        return ExcelSheetReadRequest.from(filePath).build();
    }

    private ExcelSheetReadRequest createRequest(Path path, String sheetName) {
        return ExcelSheetReadRequest.from(path).bySheetName(sheetName).build();
    }

    private ExcelSheetReadRequest createRequest(Path path, int sheetIndex) {
        return ExcelSheetReadRequest.from(path).bySheetIndex(sheetIndex).build();
    }

    private String[][] readMatrix(T sheetReader, ExcelSheetReadRequest request) throws IOException {
        return sheetReader.readSheet(request).getMatrix();
    }

    private String[][] readDefaultMatrix(ExcelSheetReadRequest request) throws IOException {
        return readMatrix(defaultSheetReader, request);
    }

    private String[][] toMatrix(List<List<String>> rows) {
        return rows.stream()
                .map(row -> row.toArray(new String[0]))
                .toArray(String[][]::new);
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class GeneralTests {

        @Test
        public void basicRead() throws IOException {
            Path testFile = TestResourceUtil.getResourceFilePath("test_data.xlsx");
            ExcelSheetReadRequest request = createRequest(testFile);

            SheetContent sheetContent = defaultSheetReader.readSheet(request);
            String[][] matrix = sheetContent.getMatrix();
            String[][] rowsMatrix = toMatrix(sheetContent.getRows());

            assertArrayEquals(EXPECTED_TEST_DATA, matrix, EXPECTED_SHEET_CONTENT_MESSAGE);
            assertArrayEquals(EXPECTED_TEST_DATA, rowsMatrix, EXPECTED_SHEET_CONTENT_MESSAGE);
        }

        /**
         * Ensure that the final matrix column counts
         * are the same for each row.  The last column
         * should have a value for at least 1 of the rows
         */
        @Test
        public void ensureColumnSize() throws IOException {
            ExcelSheetReadRequest request = createRequest(TEST_FILE, GROWING_COLUMN_LENGTH_SHEET);

            List<SheetContent> sheetContentList = defaultSheetReader.readSheets(request);
            String[][] matrix = sheetContentList.get(0).getMatrix();

            for (String[] rowValues : matrix) {
                assertEquals(3, rowValues.length);
            }
        }

        @Test
        public void readBlankSheet() throws IOException {
            ExcelSheetReadRequest request = createRequest(TEST_FILE, "BlankSheet");
            String[][] matrix = readDefaultMatrix(request);
            assertEquals(0, matrix.length, EXPECTED_ROW_COUNT_MESSAGE);
        }

        /**
         * Read Sheet that contains hidden blank rows where
         * row.getLastCellNum() == -1
         * Ensure sheet is still read correctly when these rows are encountered.
         */
        @Test
        public void negativeRowCellNumberRepro() throws IOException {
            ExcelSheetReadRequest request = createRequest(TEST_FILE, "BadRow");
            String[][] matrix = readDefaultMatrix(request);

            assertEquals("aaa", matrix[0][0]);
            assertEquals("bbb", matrix[0][1]);
            assertEquals("ccc", matrix[2][0]);
            assertEquals("ddd", matrix[2][1]);
        }

        /**
         * Sheet where the last column is filled with different
         * whitespace characters.  Since we have default values of:
         * sanitizeSpaces=true AND trimStringValues=true
         * Then this last column should be removed from the result.
         */
        @Test
        public void handleExtraBlankWhitespaceColumn() throws IOException {
            ExcelSheetReadRequest request = createRequest(TEST_FILE, "LastColWhitespace");
            String[][] matrix = readDefaultMatrix(request);
            assertEquals(1, matrix[0].length, "mismatch of expected number of columns in csv output.");
        }

        @Test
        public void readMultipleSheets() throws IOException {
            List<String> sheetNames = List.of(GROWING_COLUMN_LENGTH_SHEET, WITH_TWO_BLANK_COLUMNS_SHEET, WITH_THREE_BLANK_ROWS_SHEET);
            List<String> expectedFirstCells = List.of("aa", "aa11", "Name");

            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(TEST_FILE).bySheetNames(sheetNames).build();

            List<SheetContent> sheetContentList = defaultSheetReader.readSheets(request);
            assertEquals(sheetNames.size(), sheetContentList.size(), "Expected 3 sheets");

            for (int i = 0; i < sheetContentList.size(); i++) {
                assertEquals(expectedFirstCells.get(i), sheetContentList.get(i).getMatrix()[0][0]);
            }
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class TrimStringValuesTests {
        @Test
        public void trimStringValuesEnabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            ExcelSheetReadRequest request = createRequest(testFile);

            T sheetReader = createBuilder().trimStringValues(true).build();
            String[][] matrix = readMatrix(sheetReader, request);
            assertEquals(testValue.trim(), matrix[0][0]);
        }

        @Test
        public void trimStringValuesDisabled(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            ExcelSheetReadRequest request = createRequest(testFile);

            T sheetReader = createBuilder().trimStringValues(false).build();
            String[][] matrix = readMatrix(sheetReader, request);
            assertEquals(testValue, matrix[0][0]);
        }

        @Test
        public void trimStringValuesDefault(@TempDir Path tempDir) throws IOException {
            String testValue = "  aa bb  ";
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, testValue);
            ExcelSheetReadRequest request = createRequest(testFile);

            String[][] matrix = readDefaultMatrix(request);
            assertEquals(testValue.trim(), matrix[0][0]);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankRowTests {
        /**
         * Compare the diff of row counts with
         * skipBlankRows = true vs. false
         */
        @Test
        public void skipBlankRows() throws IOException {
            T retainBlankRowsReader = createBuilder().skipBlankRows(false).build();
            T skipBlankRowsReader = createBuilder().skipBlankRows(true).build();

            ExcelSheetReadRequest request = createRequest(TEST_FILE, WITH_THREE_BLANK_ROWS_SHEET);

            String[][] retainBlankRowsMatrix = readMatrix(retainBlankRowsReader, request);
            String[][] skipBlankRowsMatrix = readMatrix(skipBlankRowsReader, request);
            assertEquals(3, retainBlankRowsMatrix.length - skipBlankRowsMatrix.length, EXPECTED_ROW_COUNT_MESSAGE);
        }

        @Test
        public void defaultRetainBlankRows() throws IOException {
            T retainBlankRowsReader = createBuilder().skipBlankRows(false).build();

            ExcelSheetReadRequest request = createRequest(TEST_FILE, WITH_THREE_BLANK_ROWS_SHEET);

            String[][] retainBlankRowsMatrix = readMatrix(retainBlankRowsReader, request);
            String[][] defaultMatrix = readDefaultMatrix(request);
            assertEquals(retainBlankRowsMatrix.length, defaultMatrix.length, EXPECTED_ROW_COUNT_MESSAGE);
        }

        /**
         * Regardless of config settings always remove trailing blank rows
         * (i.e. the last row should contain some values for a non-blank sheet)
         */
        @Test
        public void pruneExtraBlankRows() throws IOException {
            T sheetReader = createBuilder().skipBlankRows(true).build();

            ExcelSheetReadRequest request = createRequest(TEST_FILE, "ExtraBlankRowsAfterData");
            String[][] matrix = readMatrix(sheetReader, request);
            assertEquals(2, matrix.length, EXPECTED_ROW_COUNT_MESSAGE);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class BlankColumnTests {
        /**
         * Compare the diff of column counts with
         * skipBlankColumns = true vs. false
         */
        @Test
        public void skipBlankColumns() throws IOException {
            T retainBlankColumnsReader = createBuilder().skipBlankColumns(false).build();
            T skipBlankColumnsReader = createBuilder().skipBlankColumns(true).build();

            ExcelSheetReadRequest request = createRequest(TEST_FILE, WITH_TWO_BLANK_COLUMNS_SHEET);

            String[][] retainBlankColumnsMatrix = readMatrix(retainBlankColumnsReader, request);
            String[][] skipBlankColumnsMatrix = readMatrix(skipBlankColumnsReader, request);
            assertEquals(2, retainBlankColumnsMatrix[0].length - skipBlankColumnsMatrix[0].length, EXPECTED_COLUMN_COUNT_MESSAGE);
        }

        @Test
        public void defaultRetainBlankColumns() throws IOException {
            T retainBlankColumnsReader = createBuilder().skipBlankColumns(false).build();

            ExcelSheetReadRequest request = createRequest(TEST_FILE, WITH_TWO_BLANK_COLUMNS_SHEET);

            String[][] retainBlankColumnsMatrix = readMatrix(retainBlankColumnsReader, request);
            String[][] defaultMatrix = readDefaultMatrix(request);
            assertEquals(retainBlankColumnsMatrix[0].length, defaultMatrix[0].length, EXPECTED_COLUMN_COUNT_MESSAGE);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class HiddenCellsTests {
        private static final String HIDDEN_CELLS_DATA_FILE = "skipHiddenTestData.xlsx";
        private final Path HIDDEN_CELLS_FILE = TestResourceUtil.getResourceFilePath(HIDDEN_CELLS_DATA_FILE);
        private static final String INPUT_DATA_SHEET_SUFFIX = "_Data";
        private static final String EXPECTED_DATA_SHEET_SUFFIX = "_Expected";

        /**
         * The HiddenCellsDataFile contains multiple sheets in set of 2.
         * (TestDataSheet, ExpectedResultsDataSheet)
         * where the ExpectedResultsDataSheet represents what the
         * data should look like if it were to skip/ignore the 'hidden data'
         * from the TestDataSheet
         *
         * @param sheetNamePrefix sheetPrefix
         */
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

            T sheetReader = createBuilder().skipHiddenCells(true).build();

            ExcelSheetReadRequest request = createRequest(HIDDEN_CELLS_FILE, testDataSheetName);
            ExcelSheetReadRequest expectedRequest = createRequest(HIDDEN_CELLS_FILE, expectedDataSheetName);

            String[][] actualMatrix = readMatrix(sheetReader, request);
            if (actualMatrix.length > 0) {
                assertTrue(actualMatrix[0].length > 0, "Matrix return rows with zero-length arrays");
            }

            String[][] expectedMatrix = readDefaultMatrix(expectedRequest);
            assertArrayEquals(expectedMatrix, actualMatrix);
        }

        @Test
        public void withHiddenConfigsReadAllSheets() throws IOException {
            T sheetReader = createBuilder().skipHiddenCells(true).build();

            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(HIDDEN_CELLS_FILE).allSheets().build();

            List<SheetContent> allSheetContentList = sheetReader.readSheets(request);

            // todo: update to be more resilient of sheets are in unexpected order.
            for (int i = 0; i < allSheetContentList.size(); i += 2) {
                String[][] actualMatrix = allSheetContentList.get(i).getMatrix();
                String[][] expectedMatrix = allSheetContentList.get(i + 1).getMatrix();
                assertArrayEquals(expectedMatrix, actualMatrix);
            }
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
                arguments(named("facade-diacritics", BASIC_DIACRITICS), "Façade", "Facade", false)
        );

        @ParameterizedTest
        @FieldSource("sanitizeCases")
        public void sanitizeSheet(
                SanitizeType type,
                String originalValue,
                String sanitizedValue,
                boolean isDefaultEnabled,
                @TempDir Path tempDir) throws IOException {
            Path testFile = TestExcelFileSheetUtils.createSingleCellExcelFile(tempDir, originalValue);
            ExcelSheetReadRequest request = createRequest(testFile);

            String[][] enabledMatrix = readMatrix(createSanitizeSheetReader(type, true), request);
            String[][] disabledMatrix = readMatrix(createSanitizeSheetReader(type, false), request);
            String[][] defaultMatrix = readDefaultMatrix(request);

            assertEquals(sanitizedValue, enabledMatrix[0][0]);
            assertEquals(originalValue, disabledMatrix[0][0]);

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
     * By default, the DataFormatter uses 'false' for 1904DateWindowing.
     * And this is often correct for most files.  However, if encounter
     * an Excel workbook where 1904DateWindowing is true, then the
     * default behavior will cause dates to be formatted incorrectly.
     * This test is to verify bug fix where the date values
     * would be incorrect for the 'AdvancedExcelSheetReader' class.
     */
    @Test
    public void test1904Windowing() throws IOException {
        Path date1904File = TestResourceUtil.getResourceFilePath("date1904.xlsx");

        T sheetReader = createBuilder().build();

        ExcelSheetReadRequest request = createRequest(date1904File);
        String[][] matrix = readMatrix(sheetReader, request);

        assertEquals("Due Date", matrix[0][1]);
        assertEquals("TBD", matrix[1][1]);
        assertEquals("3/15/08", matrix[2][1]);
        assertEquals("4/30/08", matrix[3][1]);
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
        public void nullRequestParameter() {
            assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                sheetReader.readSheet(null);
            });
        }

        @Test
        public void negativeSheetIndex() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                ExcelSheetReadRequest request = createRequest(Path.of("fake.xlsx"), -2);
                sheetReader.readSheet(request);
            });
            assertEquals("Indexes cannot contain negative values", thrown.getMessage());
        }

        @Test
        public void nullSheetName() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                ExcelSheetReadRequest request = createRequest(Path.of("fake.xlsx"), null);
                sheetReader.readSheet(request);
            });
            assertEquals("Names cannot contain null values", thrown.getMessage());
        }

        @Test
        public void emptySheetName() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                ExcelSheetReadRequest request = createRequest(TEST_FILE, "");
                sheetReader.readSheet(request);
            });
            assertEquals("Names cannot contain empty values", thrown.getMessage());
        }

        @Test
        public void unknownSheetName() {
            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                ExcelSheetReadRequest request = createRequest(TEST_FILE, "UnknownSheetName");
                sheetReader.readSheet(request);
            });
            assertEquals("Requested Excel sheet not found: 'UnknownSheetName'", thrown.getMessage());
        }

        @Test
        public void requestManySheetsOnSingleSheetCall() throws IOException {
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(TEST_FILE).bySheetIndex(0, 1).build();

            Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
                T sheetReader = createBuilder().build();
                sheetReader.readSheet(request);
            });
            assertEquals("Expected exactly one sheet but found 2", thrown.getMessage());
        }
    }

    @Test
    public void corruptInputFile(@TempDir Path tempDir) throws IOException {
        Path badFilePath = tempDir.resolve("bad.xlsx");
        Files.writeString(badFilePath, "bad");

        assertThrows(IOException.class, () -> {
            T sheetReader = createBuilder().build();
            ExcelSheetReadRequest request = createRequest(badFilePath, 0);
            sheetReader.readSheet(request);
        });
    }

    @ParameterizedTest
    @MethodSource("invalidInputPaths")
    public void invalidInputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws IOException {
        Path path = location != null ? Paths.get(location) : null;
        T sheetReader = createBuilder().build();

        Executable executable = () -> {
            ExcelSheetReadRequest request = createRequest(path);
            sheetReader.readSheet(request);
        };
        assertExecutableException(executable, expectedException, expectedMessage);
    }

    @ParameterizedTest
    @MethodSource("invalidInputUrls")
    public void invalidInputUrlParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws IOException {
        URL url = location != null ? new URL(location) : null;
        T sheetReader = createBuilder().build();

        Executable executable = () -> {
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(url).build();
            sheetReader.readSheet(request);
        };
        assertExecutableException(executable, expectedException, expectedMessage);
    }

    // todo - consolidate password tests
    private static final Path VALID_TEST_INPUT_PSWD_PATH = getResourceFilePath("test_data_w_pswd_1234.xlsx");

    @Test
    public void testReadPasswordProtectedFile() throws Exception {
        URL resourceUrl = getResourceFileUrl("test_data_w_pswd_1234.xlsx");

        T sheetReader = createBuilder().build();
        ExcelSheetReadRequest request = ExcelSheetReadRequest
                .from(resourceUrl)
                .password("1234")
                .build();

        String[][] matrix = readMatrix(sheetReader, request);
        assertEquals("aaa", matrix[0][0]);
        assertEquals("bbb", matrix[0][1]);
    }

    @Nested
    @DisplayName("Invalid Password Tests")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class InvalidPasswordTests {
        @ParameterizedTest
        @CsvSource({
                "bad_password, Password incorrect",
                ",no password was supplied",  // first param is null password
                "'',no password was supplied" // first param is empty string password
        })
        public void invalidPasswordCheck(String password, String expectedMessageSubstring) {
            T sheetReader = createBuilder().build();
            ExcelSheetReadRequest request = ExcelSheetReadRequest
                    .from(VALID_TEST_INPUT_PSWD_PATH)
                    .password(password)
                    .build();

            Exception exception = assertThrows(Exception.class, () -> {
                sheetReader.readSheet(request);
            });
            assertContains(expectedMessageSubstring, exception.getMessage());
        }

        private void assertContains(String substring, String value) {
            assertTrue(value.contains(substring),
                    String.format("Expected to find substring '%s' in string '%s'.", substring, value));
        }
    }

    /**
     * Junit run the executable and check that it throws the expected exception and msg.
     *
     * @param executable        executable
     * @param expectedException expectedException
     * @param expectedMessage   expectedMessage
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

    private static final String[][] EXPECTED_TEST_DATA = new String[][]{
            {"col1", "col_2", "COL3", "Col4", "Col5"},
            {"Text", "Space Text", "$_() ' extra chars", "92\"82_cell_w_quote", "extra"},
            {"11", "22", "33", "66", "row_contains_formulas"},
            {"with_dates_row", "9/8/20", "9/9/20", "1999-07-04", ""},
            {"sparse_column_row", "", "", "val", ""},
            {"", "11:33", "11:33", "", ""},
            {"", "", "", "", ""},
            {"", "", "", "", ""},
            {"misc_string-row", "value     with     spaces", "funky\"quote", "value\twith\ttabs", ""},
            {"misc_number_row", "-1", "4.765400", "1.23", ""},
            {"invalid_formula_row", "#VALUE!", "#VALUE!", "#NAME?", "i"},
            {"numeric_row", "1.888888889", "3.14", "-0.57", "1.1"}
    };
}
