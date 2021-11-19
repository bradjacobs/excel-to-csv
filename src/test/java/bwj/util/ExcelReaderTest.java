/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util;

import bwj.util.excel.ExcelReader;
import bwj.util.excel.QuoteMode;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;
import java.util.zip.GZIPInputStream;

import static org.testng.Assert.assertEquals;
import static org.testng.Assert.assertFalse;
import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.assertTrue;


public class ExcelReaderTest
{
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


//    @Test
//    public void testFoobar() throws Exception
//    {
//        String s1 = "file:///Users/bradjacobs/git/bradjacobs/excel-to-csv/src/test/resources/expected_always.csv";
//        String s2 = "file:/Users/bradjacobs/git/bradjacobs/excel-to-csv/target/test-classes/test_data.xlsx";
//        String s3 = "http://www.googlez.com";
//
//
//        URL url1 = new URL(s1);
//        URL url2 = new URL(s2);
//        URL url3 = new URL(s3);
//
//        String path2 = url2.getPath();
//        String path3 = url3.getPath();
//
//
//        int kjkjk = 00;
//
//    }

//    @Test
//    public void testExp() throws Exception
//    {
////        String path = "https://download.microsoft.com/download/1/4/E/14EDED28-6C58-4055-A65C-23B4DA81C4DE/Financial%20Sample.xlsx";
//        String path = "http://www.ilo.org/ilostat-files/Documents/ISIC.xlsx";
//        URL url = new URL(path);
//
//        URLConnection connection = url.openConnection();
//
//        Map<String, List<String>> reqProperties = connection.getRequestProperties();
//
//        // http://testws.galileo.com/gwssample/help/gwshelp/gzip_java_request_response_unannotated.htm
//        //           urlConn.setRequestProperty("Accept-Encoding","gzip");
//        connection.setConnectTimeout(20000);
//        connection.setReadTimeout(20000);
//        connection.setRequestProperty("Accept-Encoding","gzip");
//
//        String encoding1 = connection.getContentEncoding();
//
//
//        InputStream is = null;
//
//        if (encoding1 != null && encoding1.equalsIgnoreCase("gzip"))
//        {
//            is = new GZIPInputStream(connection.getInputStream());
//            System.out.println("UPDATE IS");
//        }
//        else {
//            is = connection.getInputStream();
//        }
//
//        String encoding2 = connection.getContentEncoding();
//
//        byte[] byteArray = null;
//
//        try {
//            int nRead;
//            ByteArrayOutputStream buffer = new ByteArrayOutputStream();
//            byte[] data = new byte[4096];
//
//            while ((nRead = is.read(data, 0, data.length)) != -1) {
//                buffer.write(data, 0, nRead);
//            }
//
//            byteArray = buffer.toByteArray();
//        }
//        finally {
//            is.close();
//        }
//
//        int expected = 109728;
//
//        int kjkj = 33333;
//
//    }


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
        File inputFile = getTestFileObject();

        ExcelReader excelReader = ExcelReader.builder().setQuoteMode(quoteMode).setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(inputFile);

        assertEquals(csvText, expectedCsvText, "mismatch of expected csv output");
    }

    @Test
    public void testLocalFileInputAsUrl() throws Exception
    {
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);

        URL inputFileUrl = getTestLocalFileUrl();
        ExcelReader excelReader = ExcelReader.builder().setSkipEmptyRows(false).build();
        String csvText = excelReader.convertToCsvText(inputFileUrl);
        assertEquals(csvText, expectedCsvText, "mismatch of expected csv output");
    }

    @Test
    public void testMatrixOutput() throws Exception
    {
        ExcelReader excelReader = ExcelReader.builder().build();
        File inputFile = getTestFileObject();

        String[][] csvData = excelReader.convertToDataMatrix(inputFile);
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
        File inputFile = getTestFileObject();
        String[][] csvDataOriginal = excelReader1.convertToDataMatrix(inputFile);

        ExcelReader excelReader2 = ExcelReader.builder().setSkipEmptyRows(true).build();
        String[][] csvDataSkipEmptyRows = excelReader2.convertToDataMatrix(inputFile);

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
        File inputFile = getTestFileObject();
        excelReader.convertToCsvFile(inputFile, TEST_OUTPUT_FILE);

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
        File inputFile = getTestFileObject();

        String csvText = excelReader.convertToCsvText(inputFile);
        String expectedCsvText = readResourceFile(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(csvText, expectedCsvText, "mismatch of content of saved csv file");
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
