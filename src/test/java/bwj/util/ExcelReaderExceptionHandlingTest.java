/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util;

import bwj.util.excel.ExcelReader;
import org.apache.commons.io.IOUtils;
import org.testng.annotations.Test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

import static org.testng.Assert.assertNotNull;


public class ExcelReaderExceptionHandlingTest
{
    private static final String TEST_DATA_FILE = "test_data.xlsx";


    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "Must provide a valid inputStream.")
    public void testEmptyInputStream() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().build();
        InputStream inputStream = null;
        String[][] csvData = excelReader.createCsvMatrix(inputStream);
    }

    // give an InputStream that's already been read.... any exception is fine
    @Test(expectedExceptions = { Exception.class })
    public void testInvalidInputStream() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().build();

        try (InputStream inputStream = getTestDataInputStream())
        {
            byte[] byteArray = IOUtils.toByteArray(inputStream);
            String csvText = excelReader.createCsvText(inputStream);
        }
    }

    // give an InputStream that's already closed
    @Test(expectedExceptions = { IOException.class },
        expectedExceptionsMessageRegExp = "Stream closed")
    public void testClosedInputStream() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().build();
        InputStream inputStream = getTestDataInputStream();
        try {
            byte[] byteArray = IOUtils.toByteArray(inputStream);
        }
        finally {
            inputStream.close();
        }

        String csvText = excelReader.createCsvText(inputStream);
    }

    // give an InputStream is _NOT_ an Excel file
    @Test(expectedExceptions = { IOException.class } )
    public void testNotExcelFile() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().build();
        URL url = this.getClass().getClassLoader().getResource("expected_normal.csv");
        assertNotNull(url, "unable to open test file");

        try (InputStream inputStream = url.openStream())
        {
            String csvText = excelReader.createCsvText(inputStream);
        }
    }

    @Test(expectedExceptions = { FileNotFoundException.class })
    public void testInvalidFilePath() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().build();
        String csvText = excelReader.createCsvText("/bogus/path/here/file.xlsx");
    }


    @Test(expectedExceptions = { IllegalArgumentException.class })
    public void testInvalidSheetIndex() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().setSheetIndex(99).build();

        try (InputStream inputStream = getTestDataInputStream())
        {
            String csvText = excelReader.createCsvText(inputStream);
        }
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "Unable to find sheet with name: FAKE_WORKSHEET_NAME")
    public void testInvalidSheetName() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().setSheetName("FAKE_WORKSHEET_NAME").build();

        try (InputStream inputStream = getTestDataInputStream())
        {
            String csvText = excelReader.createCsvText(inputStream);
        }
    }


    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "Cannot set quoteMode to null")
    public void testUnsetQuoteMode() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().setQuoteMode(null).build();
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "SheetIndex cannot be negative")
    public void testNegativeIndex() throws Exception {

        ExcelReader excelReader = ExcelReader.builder().setSheetIndex(-5).build();
    }




    private InputStream getTestDataInputStream()
    {
        try {
            URL resource = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
            assertNotNull(resource);
            return resource.openStream();
        }
        catch (Exception e) {
            throw new RuntimeException(String.format("Unable to read test resource file: %s.  Reason: %s", TEST_DATA_FILE, e.getMessage()), e);
        }
    }

}
