/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.testng.annotations.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.UnknownHostException;

import static org.testng.Assert.assertNotNull;

public class ExcelReaderExceptionHandlingTest {
    private static final String TEST_DATA_FILE = "test_data.xlsx";
    private static final String PSWD_DATA_FILE = "test_data_w_pswd_1234.xlsx";

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

    @Test(expectedExceptions = { UnknownHostException.class })
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

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "The outputFile cannot be an existing directory.")
    public void testSaveCsvOutputFileIsDirectory() throws Exception {
        File inputFile = getTestFileObject();

        // create a "File" that is actually pointing to an existing directory
        File directory = new File(inputFile.getParent());
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvFile(inputFile, directory);  // directory is invalid parameter
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "Illegal outputFile extension.*")
    public void testSaveCsvOutputFileInvalidExtension() throws Exception {
        File inputFile = getTestFileObject();
        File outFile = new File("outfile.exe");
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvFile(inputFile, outFile);
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = ".*outputFile path contains an illegal character.*")
    public void testSaveCsvOutputFileIllegalNullCharInPath() throws Exception {
        File inputFile = getTestFileObject();
        File outFile = new File("aaaa_||._\u0000_bbb.csv");
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvFile(inputFile, outFile);
    }

    @Test(expectedExceptions = { IllegalArgumentException.class },
            expectedExceptionsMessageRegExp = "Attempted to save CSV output file in a non-existent directory.*")
    public void testSaveCsvInvalidDirectory() throws Exception {
        File inputFile = getTestFileObject();
       File outFile = new File("/fakedirectory/myOutputFile.csv");
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvFile(inputFile, outFile);
    }

    @Test(expectedExceptions = { Exception.class },
            expectedExceptionsMessageRegExp = ".*no password was supplied.*")
    public void testMissingRequiredPassword() throws Exception {
        File inputFile = getPasswordTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().build();
        excelReader.convertToCsvFile(inputFile, new File(""));
    }

    @Test(expectedExceptions = { Exception.class },
            expectedExceptionsMessageRegExp = ".*no password was supplied.*")
    public void testBlankRequiredPassword() throws Exception {
        File inputFile = getPasswordTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().setPassword("").build();
        excelReader.convertToCsvFile(inputFile, new File(""));
    }

    @Test(expectedExceptions = { Exception.class },
            expectedExceptionsMessageRegExp = "Password incorrect")
    public void testInvalidPassword() throws Exception {
        File inputFile = getPasswordTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().setPassword("bad_password").build();
        excelReader.convertToCsvFile(inputFile, new File(""));
    }

    private File getTestFileObject() {
        URL resourceUrl = this.getClass().getClassLoader().getResource(TEST_DATA_FILE);
        assertNotNull(resourceUrl);
        return new File( resourceUrl.getPath() );
    }

    private File getPasswordTestFileObject() {
        URL resourceUrl = this.getClass().getClassLoader().getResource(PSWD_DATA_FILE);
        assertNotNull(resourceUrl);
        return new File( resourceUrl.getPath() );
    }
}
