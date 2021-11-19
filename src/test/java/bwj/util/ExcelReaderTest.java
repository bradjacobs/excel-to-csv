/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util;

import bwj.util.excel.ExcelReader;
import bwj.util.excel.QuoteMode;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;

import static org.testng.Assert.assertEquals;
import static org.testng.Assert.assertFalse;
import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.assertTrue;


public class ExcelReaderTest
{
    // can flip this boolean for extra internet tests, but it's technically not "unit testing" at that point.
    //   maybe later do the right thing and 'mock' the internet call if/when time allows.
    private static final boolean ALLOW_INTERNET_CALL = false;
    private static final String SAMPLE_INTERNET_EXCEL_FILE = "https://download.microsoft.com/download/1/4/E/14EDED28-6C58-4055-A65C-23B4DA81C4DE/Financial%20Sample.xlsx";


    private static final String TEST_DATA_FILE = "test_data.xlsx";
    private static final String TEST_SHEET_NAME = "TEST_SHEET";

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
    public void testExpectedQuoteText(QuoteMode quoteMode, String expectedResultFileName) throws Exception
    {
        String expectedCsvText = readResourceFile(expectedResultFileName);
        ExcelReader excelReader = ExcelReader.builder().setQuoteMode(quoteMode).setSkipEmptyRows(false).build();
        InputStream inputStream = getTestDataInputStream();

        String csvText = excelReader.convertToCsvText(inputStream);

        assertEquals(csvText, expectedCsvText, "mismatch of expected csv output");
    }

    @Test
    public void testMatrixOutput() throws Exception
    {
        ExcelReader excelReader = ExcelReader.builder().build();
        InputStream inputStream = getTestDataInputStream();

        String[][] csvData = excelReader.createDataMatrix(inputStream);
        assertNotNull(csvData, "expected non-null data");
        assertEquals(csvData.length, EXPECTED_ROW_COUNT, "mismatch expected number of rows");

        for (String[] rowData : csvData)
        {
            assertNotNull(rowData, "unexpected null row");
            assertEquals(rowData.length, EXPECTED_COL_COUNT, "mismatch expected columns");
        }
    }

    @Test
    public void testSkipBlankRows() throws Exception
    {
        ExcelReader excelReader1 = ExcelReader.builder().build();
        InputStream inputStream1 = getTestDataInputStream();
        String[][] csvDataOriginal = excelReader1.createDataMatrix(inputStream1);

        ExcelReader excelReader2 = ExcelReader.builder().setSkipEmptyRows(true).build();
        InputStream inputStream2 = getTestDataInputStream();
        String[][] csvDataSkipEmptyRows = excelReader2.createDataMatrix(inputStream2);

        // based on the fact we 'know' what the test data file looks like and
        // there are currently 2 empty rows, there should be a row count difference of 2.

        int originalRowCount = csvDataOriginal.length;
        int skipEmptyRowCount = csvDataSkipEmptyRows.length;

        assertEquals(skipEmptyRowCount, originalRowCount - 2, "mismatch of expected number fo rows when skipping blank rows");
    }

    @Test
    public void testSaveFile() throws Exception
    {
        cleanupTestFile(); // first check if residual file still around from a previous test.

        ExcelReader excelReader = ExcelReader.builder().build();
        InputStream inputStream = getTestDataInputStream();
        excelReader.convertToCsvFile(inputStream, TEST_OUTPUT_FILE);

        assertTrue(TEST_OUTPUT_FILE.exists(), "expected csv file was NOT created");

        String outputFileContent = new String ( Files.readAllBytes( Paths.get(TEST_OUTPUT_FILE.getAbsolutePath()) ) );

        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertEquals(outputFileContent, expectedCsvText, "mismatch of content of saved csv file");
    }


    @Test()
    public void testFilePath() throws Exception
    {
        URL fileUrl = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
        assertNotNull(fileUrl);

        ExcelReader excelReader = ExcelReader.builder().build();

        String csvText = excelReader.convertToCsvText(fileUrl);
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(csvText, expectedCsvText, "mismatch of content of saved csv file");
    }

    @Test()
    public void testReadSheetByName() throws Exception
    {
        ExcelReader excelReader = ExcelReader.builder().setSheetName(TEST_SHEET_NAME).build();
        InputStream inputStream = getTestDataInputStream();

        String csvText = excelReader.convertToCsvText(inputStream);
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(csvText, expectedCsvText, "mismatch of content of saved csv file");
    }


    @Test(enabled = ALLOW_INTERNET_CALL)
    public void testDownloadFromInternet() throws Exception
    {
        ExcelReader excelReader = ExcelReader.builder().build();
        String csvText = excelReader.convertToCsvText(new URL(SAMPLE_INTERNET_EXCEL_FILE));
        assertNotNull(csvText);
        assertTrue(csvText.length() > 0);
    }


    @AfterTest
    private void cleanupTestFile() {
        if (TEST_OUTPUT_FILE.exists()) {
            boolean wasDeleted = TEST_OUTPUT_FILE.delete();
            assertTrue(wasDeleted, "unable to delete file: " + TEST_OUTPUT_FILE.getAbsolutePath());
            assertFalse(TEST_OUTPUT_FILE.exists(), "unable to delete file");  // paranoia double-check
        }
    }


    private InputStream getTestDataInputStream() {
        try {
            URL resource = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
            assertNotNull(resource);
            return resource.openStream();
        }
        catch (Exception e) {
            throw new RuntimeException(String.format("Unable to read test resource file: %s.  Reason: %s", TEST_DATA_FILE, e.getMessage()), e);
        }
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
