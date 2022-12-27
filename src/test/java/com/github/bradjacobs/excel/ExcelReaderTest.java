/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.testng.Assert.assertEquals;
import static org.testng.Assert.assertFalse;
import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.assertTrue;

public class ExcelReaderTest
{
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

    @DataProvider(name = "quote_variations")
    public Object[][] invalidTimeZoneParams(){
        return new Object[][] {
            {QuoteMode.NORMAL, EXPECTED_NORMAL_CSV_FILE},
            {QuoteMode.ALWAYS, EXPECTED_ALWAYS_CSV_FILE},
            {QuoteMode.LENIENT, EXPECTED_LENIENT_CSV_FILE},
            {QuoteMode.NEVER, EXPECTED_NEVER_CSV_FILE},
        };
    }

    @Test(dataProvider = "quote_variations")
    public void testExpectedQuoteText(QuoteMode quoteMode, String expectedResultFileName) throws Exception {
        String expectedCsvText = readResourceFile(expectedResultFileName);
        File inputFile = getTestFileObject();

        ExcelReader excelReader = ExcelReader.builder().setQuoteMode(quoteMode).setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(inputFile);

        assertEquals(csvText, expectedCsvText, "mismatch of expected csv output");
    }

    @Test
    public void testLocalFileInputAsUrl() throws Exception {
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);

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
        assertEquals(csvData.length, EXPECTED_ROW_COUNT, "mismatch expected number of rows");

        for (String[] rowData : csvData) {
            assertNotNull(rowData, "unexpected null row");
            assertEquals(rowData.length, EXPECTED_COL_COUNT, "mismatch expected columns");
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
        assertEquals(hasEmptyRowCount-2, noEmptyRowCount, "mismatch of expected number fo rows when skipping blank rows");
    }

    @Test
    public void testSaveFile() throws Exception {
        cleanupTestFile(); // first check if residual file still around from a previous test.

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        File inputFile = getTestFileObject();
        excelReader.convertToCsvFile(inputFile, TEST_OUTPUT_FILE);
        assertTrue(TEST_OUTPUT_FILE.exists(), "expected csv file was NOT created");

        String outputFileContent = new String ( Files.readAllBytes( Paths.get(TEST_OUTPUT_FILE.getAbsolutePath()) ) );
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertEquals(outputFileContent, expectedCsvText, "mismatch of content of saved csv file");
    }

    @Test()
    public void testFilePathAsUrl() throws Exception {
        URL fileUrl = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
        assertNotNull(fileUrl);

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(fileUrl);
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(csvText, expectedCsvText, "mismatch of content of saved csv file");
    }

    @Test()
    public void testReadSheetByName() throws Exception {
        ExcelReader excelReader = ExcelReader.builder()
                .setSkipEmptyRows(false)
                .setSheetName(TEST_SHEET_NAME)
                .build();
        File inputFile = getTestFileObject();

        String csvText = excelReader.convertToCsvText(inputFile);
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(csvText, expectedCsvText, "mismatch of content of saved csv file");
    }

    @Test()
    public void testReadBlankSheet() throws Exception {
        ExcelReader excelReader = ExcelReader.builder()
                .setSheetName(TEST_BLANK_SHEET_NAME)
                .setSkipEmptyRows(false)
                .build();
        File inputFile = getTestFileObject();

        String csvText = excelReader.convertToCsvText(inputFile);
        assertNotNull(csvText);
        assertEquals(csvText, "", "expected empty csv text");
    }

    @Test
    public void testHandleExtraBlankColumns() throws Exception {
        // the 9th row (index 8) is detected with extra columns
        URL resourceUrl = this.getClass().getClassLoader().getResource("digitcodes.xlsx");
        assertNotNull(resourceUrl);

        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String[][] csvMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals(csvMatrix[0].length, 3, "mismatch of expected csv output");
    }

    // Test that some cells that only contain "whitespace" get trimmed to empty string
    @Test
    public void testTrimmingSpaces() throws Exception {
        URL resourceUrl = this.getClass().getClassLoader().getResource("spaces_data.xlsx");
        assertNotNull(resourceUrl);
        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String[][] csvMatrix = excelReader.convertToDataMatrix(resourceUrl);

        // the first row is a header row, but every other row should have
        //  and empty/blank value is the second cell
        for (int i = 1; i < csvMatrix.length; i++) {
            String[] row = csvMatrix[i];
            assertEquals(row[1], "", String.format("Expected empty string following cell '%s'", row[0]));
        }
    }

    @AfterTest
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

    private String readResourceFile(String fileName) {
        try {
            URL resource = this.getClass().getClassLoader().getResource(fileName);
            assertNotNull(resource);
            return new String ( Files.readAllBytes( Paths.get(resource.getPath()) ) );
        }
        catch (Exception e) {
            throw new RuntimeException(String.format("Unable to read test resource file: %s.  Reason: %s", fileName, e.getMessage()), e);
        }
    }
}
