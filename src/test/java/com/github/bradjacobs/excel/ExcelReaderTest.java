/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelReaderTest {
    private static final String TEST_DATA_FILE = "test_data.xlsx";
    private static final String TEST_SHEET_NAME = "TEST_SHEET";
    private static final String TEST_BLANK_SHEET_NAME = "TEST_BLANK_SHEET";

    private static final String EXPECTED_NORMAL_CSV_FILE = "expected_normal.csv";
    private static final String EXPECTED_ALWAYS_CSV_FILE = "expected_always.csv";
    private static final String EXPECTED_LENIENT_CSV_FILE = "expected_lenient.csv";
    private static final String EXPECTED_NEVER_CSV_FILE = "expected_never.csv";

    private static final int EXPECTED_ROW_COUNT = 12;
    private static final int EXPECTED_COL_COUNT = 5;

    private static final String TEST_OUTPUT_FILE_NAME = "test_output.csv";
    private static final File TEST_OUTPUT_FILE = new File(TEST_OUTPUT_FILE_NAME);

    private static Stream<Arguments> quoteVariations() {
        return Stream.of(
                Arguments.of(QuoteMode.NORMAL, EXPECTED_NORMAL_CSV_FILE),
                Arguments.of(QuoteMode.ALWAYS, EXPECTED_ALWAYS_CSV_FILE),
                Arguments.of(QuoteMode.LENIENT, EXPECTED_LENIENT_CSV_FILE),
                Arguments.of(QuoteMode.NEVER, EXPECTED_NEVER_CSV_FILE)
        );
    }

    @ParameterizedTest
    @MethodSource("quoteVariations")
    void testExpectedQuoteTextFileParam(QuoteMode quoteMode, String expectedResultFileName) throws Exception {
        String expectedCsvText = readResourceFileText(expectedResultFileName);
        File inputFile = getTestFileObject();

        ExcelReader excelReader = ExcelReader.builder().setQuoteMode(quoteMode).setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(inputFile);

        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    @ParameterizedTest
    @MethodSource("quoteVariations")
    void testExpectedQuoteTextPathParam(QuoteMode quoteMode, String expectedResultFileName) throws Exception {
        String expectedCsvText = readResourceFileText(expectedResultFileName);
        Path inputFile = getTestFileObject().toPath();

        ExcelReader excelReader = ExcelReader.builder().setQuoteMode(quoteMode).setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(inputFile);

        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    @Test
    public void testLocalFileInputAsUrl() throws Exception {
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);

        URL inputFileUrl = getTestLocalFileUrl();
        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(inputFileUrl);
        assertEquals(csvText, expectedCsvText, "mismatch of expected csv output");
    }

    @Test
    public void testMatrixOutput() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        File inputFile = getTestFileObject();

        String[][] csvData = excelReader.convertToDataMatrix(inputFile);
        assertNotNull(csvData, "expected non-null data");
        assertEquals(EXPECTED_ROW_COUNT, csvData.length,"mismatch expected number of rows");

        for (String[] rowData : csvData) {
            assertNotNull(rowData, "unexpected null row");
            assertEquals(EXPECTED_COL_COUNT, rowData.length,"mismatch expected columns");
        }
    }

    @Test
    public void testSkipBlankRows() throws Exception {
        File inputFile = getTestFileObject();

        ExcelReader excelReader1 = ExcelReader.builder()
                .setSkipEmptyRows(false)
                .build();
        String[][] csvDataHasEmptyRows = excelReader1.convertToDataMatrix(inputFile);

        ExcelReader excelReader2 = ExcelReader.builder()
                .setSkipEmptyRows(true)
                .build();
        String[][] csvDataNoEmptyRows = excelReader2.convertToDataMatrix(inputFile);

        // based on the fact we 'know' what the test data file looks like and
        // there are currently 2 empty rows, there should be a row count difference of 2.
        int hasEmptyRowCount = csvDataHasEmptyRows.length;
        int noEmptyRowCount = csvDataNoEmptyRows.length;
        assertEquals(noEmptyRowCount,hasEmptyRowCount-2, "mismatch of expected number fo rows when skipping blank rows");
    }

    @Test
    public void testSaveFile() throws Exception {
        cleanupTestFile(); // first check if residual file still around from a previous test.

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        File inputFile = getTestFileObject();
        excelReader.convertToCsvFile(inputFile, TEST_OUTPUT_FILE);
        assertTrue(TEST_OUTPUT_FILE.exists(), "expected csv file was NOT created");

        String outputFileContent =
                Files.readString(Paths.get(TEST_OUTPUT_FILE.getAbsolutePath()), StandardCharsets.UTF_8);
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
        assertEquals(expectedCsvText, outputFileContent, "mismatch of content of saved csv file");
    }

    @Test
    public void testFilePathAsUrl() throws Exception {
        URL fileUrl = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
        assertNotNull(fileUrl);

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(fileUrl);
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(expectedCsvText, csvText, "mismatch of content of saved csv file");
    }

    @Test
    public void testReadSheetByName() throws Exception {
        ExcelReader excelReader = ExcelReader.builder()
                .setSkipEmptyRows(false)
                .setSheetName(TEST_SHEET_NAME.toLowerCase())  // use lower to confirm match is case-insensitive.
                .build();
        File inputFile = getTestFileObject();

        String csvText = excelReader.convertToCsvText(inputFile);
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(expectedCsvText, csvText, "mismatch of content of saved csv file");
    }

    @Test
    public void testReadBlankSheet() throws Exception {
        ExcelReader excelReader = ExcelReader.builder()
                .setSheetName(TEST_BLANK_SHEET_NAME)
                .setSkipEmptyRows(false)
                .build();
        File inputFile = getTestFileObject();

        String csvText = excelReader.convertToCsvText(inputFile);
        assertNotNull(csvText);
        assertEquals("", csvText, "expected empty csv text");
    }

    // Test that some cells that only contain "whitespace" get trimmed to empty string
    @Test
    public void testTrimmingSpaces() throws Exception {
        URL resourceUrl = this.getClass().getClassLoader().getResource("spaces_data.xlsx");
        assertNotNull(resourceUrl);
        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String[][] csvMatrix = excelReader.convertToDataMatrix(resourceUrl);

        // the first row is a header row, but every other row should have
        //  and empty/blank value in the second cell
        for (int i = 1; i < csvMatrix.length; i++) {
            String[] row = csvMatrix[i];
            assertEquals("", row[1], String.format("Expected empty string following cell '%s'", row[0]));
        }
    }

    // read a worksheet where the last column only has whitespace.
    //   in this case the column should not be counted.
    @Test
    public void testHandleExtraBlankColumns() throws Exception {
        URL resourceUrl = this.getClass().getClassLoader().getResource("spaces_data.xlsx");
        assertNotNull(resourceUrl);

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).setSheetName("LAST_COL_WHITESPACE").build();
        String[][] csvMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals(1, csvMatrix[0].length, "mismatch of expected number of columns in csv output.");
    }

    // Happy Path testcase reading an Excel file that is password protected.
    @Test
    public void testReadPasswordProtectedFile() throws Exception {
        URL resourceUrl = this.getClass().getClassLoader().getResource("test_data_w_pswd_1234.xlsx");
        assertNotNull(resourceUrl);

        ExcelReader excelReader = ExcelReader.builder().setPassword("1234").build();
        String[][] csvMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals("aaa", csvMatrix[0][0]);
        assertEquals("bbb", csvMatrix[0][1]);
    }

    // Test special case where a row without cells could
    //   cause an ArrayIndexOutOfBoundsException
    @Test
    public void testBadRowRepro() throws Exception {
        URL resourceUrl = this.getClass().getClassLoader().getResource("repro.xlsx");
        assertNotNull(resourceUrl);

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String[][] csvMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals("aaa", csvMatrix[0][0]);
        assertEquals("bbb", csvMatrix[0][1]);
        assertEquals("ccc", csvMatrix[2][0]);
        assertEquals("ddd", csvMatrix[2][1]);
    }

    @AfterEach
    private void cleanupTestFile() {
        if (TEST_OUTPUT_FILE.exists()) {
            boolean wasDeleted = TEST_OUTPUT_FILE.delete();
            assertTrue(wasDeleted, "unable to delete file: " + TEST_OUTPUT_FILE.getAbsolutePath());
            assertFalse(TEST_OUTPUT_FILE.exists(), "unable to delete file");  // paranoia double-check
        }
    }

    private File getTestFileObject() {
        return new File( getTestLocalFileUrl().getPath() );
    }

    private URL getTestLocalFileUrl() {
        URL resourceUrl = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
        assertNotNull(resourceUrl);
        return resourceUrl;
    }

    private String readResourceFileText(String fileName) {
        try (InputStream is =  this.getClass().getClassLoader().getResourceAsStream(fileName)) {
            assertNotNull(is);
            return IOUtils.toString(is, StandardCharsets.UTF_8);
        }
        catch (IOException e) {
            throw new RuntimeException(String.format("Unable to read test resource file: %s.  Reason: %s", fileName, e.getMessage()), e);
        }
    }
}
