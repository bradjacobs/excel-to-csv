/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.net.UnknownHostException;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

// TODO - look at refactor options to minimize some of the redundant code
//   for the tests with 'File' vs 'Path'
public class ExcelReaderExceptionHandlingTest {
    private static final String TEST_DATA_FILE = "test_data.xlsx";
    private static final String PSWD_DATA_FILE = "test_data_w_pswd_1234.xlsx";

    @Test
    public void testMissingPathParam() {
        ExcelReader excelReader = ExcelReader.builder().build();
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToDataMatrix((Path) null);
        });
        assertEquals("Must provide an input file.", exception.getMessage());
    }
    @Test
    public void testMissingFileParam() {
        ExcelReader excelReader = ExcelReader.builder().build();
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToDataMatrix((File) null);
        });
        assertEquals("Must provide an input file.", exception.getMessage());
    }

    @Test
    public void testMissingUrlParam() {
        ExcelReader excelReader = ExcelReader.builder().build();
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToDataMatrix((URL) null);
        });
        assertEquals("Must provide an input url.", exception.getMessage());
    }

    // give a URL that is _NOT_ an Excel file
    @Test
    public void testNotExcelFileUrl() {
        ExcelReader excelReader = ExcelReader.builder().build();
        URL url = this.getClass().getClassLoader().getResource("expected_normal.csv");
        assertNotNull(url, "unable to open test file");
        assertThrows(IOException.class, () -> {
            excelReader.convertToCsvText(url);
        });
    }

    @Test
    public void testInvalidFileObjectPath() {
        ExcelReader excelReader = ExcelReader.builder().build();
        assertThrows(FileNotFoundException.class, () -> {
            excelReader.convertToCsvText(new File("/bogus/path/here/file.xlsx"));
        });
    }

    @Test
    public void testInvalidPathObjectPath() {
        ExcelReader excelReader = ExcelReader.builder().build();
        assertThrows(FileNotFoundException.class, () -> {
            excelReader.convertToCsvText(Paths.get("/bogus/path/here/file.xlsx"));
        });
    }

    @Test
    public void testInvalidUrlObjectPath() {
        ExcelReader excelReader = ExcelReader.builder().build();
        assertThrows(UnknownHostException.class, () -> {
            excelReader.convertToCsvText(new URL("https://www.zxfake12.com/foo/bar.html"));
        });
    }

    @Test
    public void testUnsupportedProtocolUrl() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        URL invalidUrl = new URL("jar:file:/C:/parser/jar/parser.jar!/test.xml");

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvText(invalidUrl);
        });
        assertEquals("URL has an unsupported protocol: jar", exception.getMessage());
    }

    @Test
    public void testInvalidSheetIndex() {
        ExcelReader excelReader = ExcelReader.builder().setSheetIndex(99).build();
        File inputFile = getTestFileObject();
        assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvText(inputFile);
        });
    }

    @Test
    public void testInvalidSheetName() {
        File inputFile = getTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().setSheetName("FAKE_WORKSHEET_NAME").build();
        assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvText(inputFile);
        });
    }

    @Test
    public void testUnsetQuoteMode() {
        assertThrows(IllegalArgumentException.class, () -> {
            ExcelReader.builder().setQuoteMode(null).build();
        });
    }

    @Test
    public void testNegativeIndex() {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            ExcelReader.builder().setSheetIndex(-5).build();
        });
        assertEquals("SheetIndex cannot be negative", exception.getMessage());
    }

    @Test
    public void testSaveCsvMissingOutputFileObjectParam() {
        File inputFile = getTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, null);
        });
        assertEquals("Must supply outputFile location to save CSV data.", exception.getMessage());
    }

    @Test
    public void testSaveCsvMissingOutputPathObjectParam() {
        Path inputFile = getTestFileObject().toPath();
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, null);
        });
        assertEquals("Must supply outputFile location to save CSV data.", exception.getMessage());
    }

    @Test
    public void testSaveCsvOutputFileObjectIsDirectory() {
        File inputFile = getTestFileObject();

        // create a "File" that is actually pointing to an existing directory
        File directory = new File(inputFile.getParent());
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, directory);  // directory is invalid parameter
        });
        assertEquals("The outputFile cannot be an existing directory.", exception.getMessage());
    }

    @Test
    public void testSaveCsvOutputPathObjectIsDirectory() {
        File inputFile = getTestFileObject();

        // create a "File" that is actually pointing to an existing directory
        Path directory = new File(inputFile.getParent()).toPath();
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile.toPath(), directory);  // directory is invalid parameter
        });
        assertEquals("The outputFile cannot be an existing directory.", exception.getMessage());
    }

    @Test
    public void testSaveCsvOutputFileObjectInvalidExtension() {
        File inputFile = getTestFileObject();
        File outFile = new File("outfile.exe");
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, outFile);
        });
        assertContains("Illegal outputFile extension", exception.getMessage());
    }

    @Test
    public void testSaveCsvOutputPathObjectInvalidExtension() {
        Path inputFile = getTestFileObject().toPath();
        Path outFile = new File("outfile.exe").toPath();
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, outFile);
        });
        assertContains("Illegal outputFile extension", exception.getMessage());
    }

    @Test
    public void testSaveCsvOutputFileIllegalNullCharInPath() {
        File inputFile = getTestFileObject();
        File outFile = new File("aa().?aa_||._\u0000_bbb.csv");
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, outFile);
        });
        assertContains(
                "path contains an illegal character",
                exception.getMessage());
    }

    @Test
    public void testSaveCsvInvalidDirectoryFileObject() {
        File inputFile = getTestFileObject();
        File outFile = new File("/fakedirectory/myOutputFile.csv");
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, outFile);
        });
        assertContains(
                "Attempted to save CSV output file in a non-existent directory",
                exception.getMessage());
    }

    @Test
    public void testSaveCsvInvalidDirectoryPathObject() {
        Path inputFile = getTestFileObject().toPath();
        Path outFile = new File("/fakedirectory/myOutputFile.csv").toPath();
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(inputFile, outFile);
        });
        assertContains(
                "Attempted to save CSV output file in a non-existent directory",
                exception.getMessage());
    }

    @Test
    public void testMissingRequiredPassword() {
        File inputFile = getPasswordTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(Exception.class, () -> {
            excelReader.convertToCsvText(inputFile);
        });
        assertContains("no password was supplied", exception.getMessage());
    }

    @Test
    public void testBlankRequiredPassword() {
        File inputFile = getPasswordTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().setPassword("").build();

        Exception exception = assertThrows(Exception.class, () -> {
            excelReader.convertToCsvText(inputFile);
        });
        assertContains("no password was supplied", exception.getMessage());
    }

    @Test
    public void testInvalidPassword() {
        File inputFile = getPasswordTestFileObject();
        ExcelReader excelReader = ExcelReader.builder().setPassword("bad_password").build();

        Exception exception = assertThrows(Exception.class, () -> {
            excelReader.convertToCsvText(inputFile);
        });
        assertEquals("Password incorrect", exception.getMessage());
    }

    private void assertContains(String subString, String mainString) {
        assertTrue(mainString.contains(subString),
                String.format("Expected to find substring '%s' in string '%s'.", subString, mainString));
    }

    private File getTestFileObject() {
        return getResourceFileObject(TEST_DATA_FILE);
    }

    private File getPasswordTestFileObject() {
        return getResourceFileObject(PSWD_DATA_FILE);
    }

    private File getResourceFileObject(String fileName) {
        URL resourceUrl = this.getClass().getClassLoader().getResource(fileName);
        assertNotNull(resourceUrl);
        return new File(resourceUrl.getPath());
    }
}
