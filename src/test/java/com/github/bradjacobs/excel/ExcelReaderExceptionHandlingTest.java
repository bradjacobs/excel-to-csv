/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.testng.annotations.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;

import static org.testng.Assert.assertNotNull;

public class ExcelReaderExceptionHandlingTest
{
    private static final String TEST_DATA_FILE = "test_data.xlsx";

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "Must provide an input file.")
    public void testMissingFile() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        String[][] csvData = excelReader.convertToDataMatrix((File)null);
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "Must provide an input url.")
    public void testMissingUrl() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        String[][] csvData = excelReader.convertToDataMatrix((URL)null);
    }

    // give a URL that is _NOT_ an Excel file
    @Test(expectedExceptions = { IOException.class } )
    public void testNotExcelFile() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        URL url = this.getClass().getClassLoader().getResource("expected_normal.csv");
        assertNotNull(url, "unable to open test file");
        String csvText = excelReader.convertToCsvText(url);
    }

    @Test(expectedExceptions = { FileNotFoundException.class })
    public void testInvalidFilePath() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvText(new File("/bogus/path/here/file.xlsx"));
    }

    @Test(expectedExceptions = { java.net.UnknownHostException.class })
    public void testInvalidUrlPath() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvText(new URL("https://www.zxfake12.com/foo/bar.html"));
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "URL has an unsupported protocol: jar")
    public void testUnsupportedProtocolUrl() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        URL invalidUrl = new URL("jar:file:/C:/parser/jar/parser.jar!/test.xml");
        excelReader.convertToCsvText(invalidUrl);
    }

    @Test(expectedExceptions = { IllegalArgumentException.class })
    public void testInvalidSheetIndex() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().setSheetIndex(99).build();
        File inputFile = getTestFileObject();
        String csvText = excelReader.convertToCsvText(inputFile);
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "Unable to find sheet with name: FAKE_WORKSHEET_NAME")
    public void testInvalidSheetName() throws Exception {
        File inputFile = getTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().setSheetName("FAKE_WORKSHEET_NAME").build();
        String csvText = excelReader.convertToCsvText(inputFile);
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "Cannot set quoteMode to null")
    public void testUnsetQuoteMode() {
        ExcelReader.builder().setQuoteMode(null).build();
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
        expectedExceptionsMessageRegExp = "SheetIndex cannot be negative")
    public void testNegativeIndex() {
        ExcelReader.builder().setSheetIndex(-5).build();
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "Must supply outputFile location to save CSV data.")
    public void testSaveCsvMissingOutputFile() throws Exception {
        File inputFile = getTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvFile(inputFile, null);
    }

    private File getTestFileObject() {
        URL resourceUrl = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
        assertNotNull(resourceUrl);
        return new File( resourceUrl.getPath() );
    }
}
