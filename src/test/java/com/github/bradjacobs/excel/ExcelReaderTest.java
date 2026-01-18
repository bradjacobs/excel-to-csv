/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.io.FileUtils;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.File;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileObject;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileUrl;
import static com.github.bradjacobs.excel.util.TestResourceUtil.readResourceFileText;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

// TODO - this class needs an overhaul.
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
        File inputFile = getResourceFileObject(TEST_DATA_FILE);
        ExcelReader excelReader = ExcelReader.builder().quoteMode(quoteMode).build();
        String csvText = excelReader.convertToCsvText(inputFile);

        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    @ParameterizedTest
    @MethodSource("quoteVariations")
    void testExpectedQuoteTextPathParam(QuoteMode quoteMode, String expectedResultFileName) throws Exception {
        String expectedCsvText = readResourceFileText(expectedResultFileName);
        Path inputFile = getResourceFilePath(TEST_DATA_FILE);

        ExcelReader excelReader = ExcelReader.builder().quoteMode(quoteMode).build();
        String csvText = excelReader.convertToCsvText(inputFile);

        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    @Test
    public void testLocalFileInputAsUrl() throws Exception {
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);

        URL inputFileUrl = getResourceFileUrl(TEST_DATA_FILE);
        ExcelReader excelReader = ExcelReader.builder().build();
        String csvText = excelReader.convertToCsvText(inputFileUrl);
        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    @Test
    public void testMatrixOutput() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        File inputFile = getResourceFileObject(TEST_DATA_FILE);

        String[][] csvData = excelReader.convertToDataMatrix(inputFile);
        assertNotNull(csvData, "expected non-null data");
        assertEquals(EXPECTED_ROW_COUNT, csvData.length, "mismatch expected number of rows");

        for (String[] rowData : csvData) {
            assertNotNull(rowData, "unexpected null row");
            assertEquals(EXPECTED_COL_COUNT, rowData.length, "mismatch expected columns");
        }
    }

    @Test
    public void testFilePathAsUrl() throws Exception {
        URL fileUrl = getResourceFileUrl(TEST_DATA_FILE);
        ExcelReader excelReader = ExcelReader.builder().build();
        String csvText = excelReader.convertToCsvText(fileUrl);
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(expectedCsvText, csvText, "mismatch of content of saved csv file");
    }

    @Test
    public void testReadSheetByName() throws Exception {
        ExcelReader excelReader = ExcelReader.builder()
                .sheetName(TEST_SHEET_NAME.toLowerCase()) // use lower to confirm match is case-insensitive.
                .build();
        File inputFile = getResourceFileObject(TEST_DATA_FILE);

        String csvText = excelReader.convertToCsvText(inputFile);
        String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
        assertNotNull(csvText);
        assertEquals(expectedCsvText, csvText, "mismatch of content of saved csv file");
    }

    @Test
    public void testReadBlankSheet() throws Exception {
        ExcelReader excelReader = ExcelReader.builder()
                .sheetName(TEST_BLANK_SHEET_NAME)
                .build();
        File inputFile = getResourceFileObject(TEST_DATA_FILE);

        String csvText = excelReader.convertToCsvText(inputFile);
        assertNotNull(csvText);
        assertEquals("", csvText, "expected empty csv text");
    }

    // Happy Path testcase reading an Excel file that is password protected.
    @Test
    public void testReadPasswordProtectedFile() throws Exception {
        URL resourceUrl = getResourceFileUrl("test_data_w_pswd_1234.xlsx");
        ExcelReader excelReader = ExcelReader.builder().password("1234").build();
        String[][] dataMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals("aaa", dataMatrix[0][0]);
        assertEquals("bbb", dataMatrix[0][1]);
    }

    @Test
    public void testFilterBlankCoumns() throws Exception {
        URL resourceUrl = getResourceFileUrl("repro.xlsx");
        ExcelReader excelReader = ExcelReader.builder()
                .sheetName("WithBlankColumns1")
                .removeBlankColumns(true)
                .build();
        String[][] dataMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals(3, dataMatrix[0].length, "mismatch expected column count");
    }

    @Test
    public void testFilterFirstBlankCoumns() throws Exception {
        URL resourceUrl = getResourceFileUrl("repro.xlsx");
        ExcelReader excelReader = ExcelReader.builder()
                .sheetName("WithBlankColumns2")
                .removeBlankColumns(true)
                .build();
        String[][] dataMatrix = excelReader.convertToDataMatrix(resourceUrl);
        assertEquals(1, dataMatrix[0].length, "mismatch expected column count");
    }

    @Nested
    class SavingCsvFileTests {
        @Test
        public void testSavePathObject(@TempDir Path tempDir) throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            ExcelReader excelReader = ExcelReader.builder().build();
            Path inputFile = getResourceFilePath(TEST_DATA_FILE);
            excelReader.convertToCsvFile(inputFile, testOutputFile);
            assertTrue(Files.exists(testOutputFile), "expected csv file was NOT created");

            String outputFileContent =
                    Files.readString(testOutputFile, StandardCharsets.UTF_8);
            String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
            assertEquals(expectedCsvText, outputFileContent, "mismatch of content of saved csv file");
        }

        @Test
        public void testSaveFileObject(@TempDir Path tempDir) throws Exception {
            File testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME).toFile();

            ExcelReader excelReader = ExcelReader.builder().build();
            File inputFile = getResourceFileObject(TEST_DATA_FILE);
            excelReader.convertToCsvFile(inputFile, testOutputFile);
            assertTrue(testOutputFile.exists(), "expected csv file was NOT created");

            String outputFileContent =
                    Files.readString(Paths.get(testOutputFile.getAbsolutePath()), StandardCharsets.UTF_8);
            String expectedCsvText = readResourceFileText(EXPECTED_NORMAL_CSV_FILE);
            assertEquals(expectedCsvText, outputFileContent, "mismatch of content of saved csv file");
        }

        // if the output csv file saved does NOT have any unicode,
        //  then the 'SaveUnicodeFileWithBom' flag should have no effect.
        @Test
        public void testBomFlagWithoutUnicode(@TempDir Path tempDir) throws Exception {
            File testOutputFile1 = tempDir.resolve("test_bom_flag_off.csv").toFile();
            File testOutputFile2 = tempDir.resolve("test_bom_flag_on.csv").toFile();

            File inputFile = getResourceFileObject(TEST_DATA_FILE);

            ExcelReader excelReader1 = ExcelReader.builder().saveUnicodeFileWithBom(false).build();
            excelReader1.convertToCsvFile(inputFile, testOutputFile1);
            ExcelReader excelReader2 = ExcelReader.builder().saveUnicodeFileWithBom(true).build();
            excelReader2.convertToCsvFile(inputFile, testOutputFile2);

            assertEquals(testOutputFile1.length(), testOutputFile2.length(), "expect 2 files to be the same size");
        }

        @Test
        public void testBomFlagWithUnicode(@TempDir Path tempDir) throws Exception {
            File testOutputFile1 = tempDir.resolve("test_bom_flag_off.csv").toFile();
            File testOutputFile2 = tempDir.resolve("test_bom_flag_on.csv").toFile();

            URL inputFileUrl = getResourceFileUrl("repro.xlsx");
            String sheetName = "WithUnicode";

            ExcelReader excelReader1 = ExcelReader.builder()
                    .saveUnicodeFileWithBom(false).sheetName(sheetName).build();
            ExcelReader excelReader2 = ExcelReader.builder()
                    .saveUnicodeFileWithBom(true).sheetName(sheetName).build();

            excelReader1.convertToCsvFile(inputFileUrl, testOutputFile1);
            excelReader2.convertToCsvFile(inputFileUrl, testOutputFile2);

            String csvFileString1 = FileUtils.readFileToString(testOutputFile1, StandardCharsets.UTF_8);
            String csvFileString2 = FileUtils.readFileToString(testOutputFile2, StandardCharsets.UTF_8);

            char expectedBom = '\uFEFF';
            // the file that had the bom flag turned on should have it as the first character
            assertEquals(expectedBom, csvFileString2.charAt(0), "Didn't find expected Bom on saved csv file");

            // now if remove this first bom character, then the 2 strings should be equal
            String trimmedCsvFileString2 = csvFileString2.substring(1);
            assertEquals(csvFileString1, trimmedCsvFileString2);
        }
    }
}
