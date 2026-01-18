/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.function.Executable;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.UnknownHostException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.api.TestInstance.Lifecycle;
import static org.junit.jupiter.params.provider.Arguments.arguments;

/**
 * Test Cases explicitly for bad input error handling
 * for the ExcelReader class.
 */
@TestInstance(Lifecycle.PER_CLASS)
public class ExcelReaderExceptionHandlingTest {
    private static final Path VALID_TEST_INPUT_PATH = getResourceFilePath("test_data.xlsx");
    private static final File VALID_TEST_INPUT_FILE = VALID_TEST_INPUT_PATH.toFile();
    private static final Path VALID_TEST_INPUT_PSWD_PATH = getResourceFilePath("test_data_w_pswd_1234.xlsx");
    private static final File VALID_TEST_INPUT_PSWD_FILE = VALID_TEST_INPUT_PSWD_PATH.toFile();
    private static final Path VALID_OUT_PATH = Paths.get("out.csv");
    private static final File VALID_OUT_FILE = VALID_OUT_PATH.toFile();

    private static final ExcelReader DEFAULT_EXCEL_READER = ExcelReader.builder().build();

    private static final String TEXT_FILE_PATH = getResourceFilePath("fake.txt").toAbsolutePath().toString();
    private static final String DIR_PATH = Paths.get("").toAbsolutePath().toString();

    private static final String INPUT_FILE_NOT_FOUND_PATH = "/bogus/path/here/file.xlsx";
    private static final String INPUT_FILE_NOT_FOUND_URL = "https://www.zxfake12.com/foo/bar.html";

    // DEV NOTE: the abstract class is a 'workaround' such that the '@Nested' classes can
    //     access the argument lists because Nested classes are not supposed to be 'static',
    //     but yet they need to be in order to access static methodSource.
    //     using @TestInstance(Lifecycle.PER_CLASS) seems to work, BUT it could
    //     still produce a bunch of warnings which are distracting.
    abstract static class AbstractNestedTestClass {
        protected static List<Arguments> invalidInputPaths() {
            return Arrays.asList(
                    arguments(named("File Not Found", INPUT_FILE_NOT_FOUND_PATH), FileNotFoundException.class, "Invalid Excel file path: /bogus/path/here/file.xlsx"),
                    arguments(named("Null Input", null), IllegalArgumentException.class, "Must provide an input file."),
                    arguments(named("Invalid Excel File Input", TEXT_FILE_PATH), IOException.class, null),
                    arguments(named("Directory Input", DIR_PATH), IllegalArgumentException.class, "The input file is a directory.")
            );
        }

        protected static List<Arguments> invalidInputUrls() {
            return Arrays.asList(
                    arguments(named("Url Not Found", INPUT_FILE_NOT_FOUND_URL), UnknownHostException.class, null),
                    arguments(named("Null Input", null), IllegalArgumentException.class, "Must provide an input url."),
                    arguments(named("Directory Input", "file:///"), IllegalArgumentException.class, "The input file is a directory."),
                    arguments(named("Invalid Url Protocol", "jar:file:/C:/foo/jar/parser.jar!/test.xlsx"), IllegalArgumentException.class, "URL has an unsupported protocol: jar")
            );
        }

        protected static List<Arguments> invalidOutputPaths() {
            return Arrays.asList(
                    arguments(named("Invalid Output file extension", "outfile.exe"), IllegalArgumentException.class, "Illegal outputFile extension 'exe'.  Must be either 'csv', 'txt' or blank"),
                    arguments(named("Invalid Output directory", "/fakeDirectory/myOutputFile.csv"), IllegalArgumentException.class, "Attempted to save CSV output file in a non-existent directory: /fakeDirectory/myOutputFile.csv"),
                    arguments(named("Null Csv Output Param", null), IllegalArgumentException.class, "Must supply outputFile location to save CSV data."),
                    arguments(named("Directory Output Param", DIR_PATH), IllegalArgumentException.class, "The outputFile cannot be an existing directory.")
            );
        }
    }

    @Nested
    @DisplayName("ConvertToCsvFile Invalid Param Tests")
    @TestInstance(Lifecycle.PER_CLASS)
    class ConvertToCsvFileInvalidPathTests extends AbstractNestedTestClass {
        @ParameterizedTest
        @MethodSource("invalidInputPaths")
        public void invalidInputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            Path path = location != null ? Paths.get(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvFile(path, VALID_OUT_PATH);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputPaths")
        public void invalidInputFileParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            File file = location != null ? new File(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvFile(file, VALID_OUT_FILE);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputUrls")
        public void invalidInputUrlParamWithPath(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            URL url = location != null ? new URL(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvFile(url, VALID_OUT_PATH);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputUrls")
        public void invalidInputUrlParamWithFile(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            URL url = location != null ? new URL(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvFile(url, VALID_OUT_FILE);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidOutputPaths")
        public void invalidOutputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            Path path = location != null ? Paths.get(location) : null;
            Executable executable1 = () -> DEFAULT_EXCEL_READER.convertToCsvFile(VALID_TEST_INPUT_PATH, path);
            assertExecutableException(executable1, expectedException, expectedMessage);

            // repeat the same test using URL input param.  \
            //   The invalid CSV out param should be detected _BEFORE_ the invalid url is detected
            URL url = new URL("http://somesite.com/file.xlsx");
            Executable executable2 = () -> DEFAULT_EXCEL_READER.convertToCsvFile(url, path);
            assertExecutableException(executable2, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidOutputPaths")
        public void invalidOutputFileParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            File file = location != null ? new File(location) : null;
            Executable executable1 = () -> DEFAULT_EXCEL_READER.convertToCsvFile(VALID_TEST_INPUT_FILE, file);
            assertExecutableException(executable1, expectedException, expectedMessage);

            // repeat the same test using URL input param.  \
            //   The invalid CSV out param should be detected _BEFORE_ the invalid url is detected
            URL url = new URL("http://somesite.com/file.xlsx");
            Executable executable2 = () -> DEFAULT_EXCEL_READER.convertToCsvFile(url, file);
            assertExecutableException(executable2, expectedException, expectedMessage);
        }
    }

    @Nested
    @DisplayName("ConvertToCsvText Invalid Param Tests")
    @TestInstance(Lifecycle.PER_CLASS)
    class ConvertToCsvTextStringInvalidPathTests extends AbstractNestedTestClass {
        @ParameterizedTest
        @MethodSource("invalidInputPaths")
        public void invalidInputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            Path path = location != null ? Paths.get(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvText(path);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputPaths")
        public void invalidInputFileParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            File file = location != null ? new File(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvText(file);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputUrls")
        public void invalidInputUrlParamWithPath(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            URL url = location != null ? new URL(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToCsvText(url);
            assertExecutableException(executable, expectedException, expectedMessage);
        }
    }

    @Nested
    @DisplayName("ConvertToDataMatrix Invalid Param Tests")
    @TestInstance(Lifecycle.PER_CLASS)
    class ConvertToDataMatrixInvalidPathTests extends AbstractNestedTestClass {
        @ParameterizedTest
        @MethodSource("invalidInputPaths")
        public void invalidInputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            Path path = location != null ? Paths.get(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToDataMatrix(path);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputPaths")
        public void invalidInputFileParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            File file = location != null ? new File(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToDataMatrix(file);
            assertExecutableException(executable, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("invalidInputUrls")
        public void invalidInputUrlParamWithPath(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            URL url = location != null ? new URL(location) : null;
            Executable executable = () -> DEFAULT_EXCEL_READER.convertToDataMatrix(url);
            assertExecutableException(executable, expectedException, expectedMessage);
        }
    }

    /**
     * Junit run the executable and check that it throws the expected exception and msg.
     * @param executable executable
     * @param expectedException expectedException
     * @param expectedMessage expectedMessage
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

    @Nested
    @DisplayName("Invalid Builder Param Tests")
    @TestInstance(Lifecycle.PER_CLASS)
    class InvalidBuilderParamTests  {
        @Test
        public void outOfBoundsSheetIndex() {
            ExcelReader excelReader = ExcelReader.builder().sheetIndex(99).build();
            assertThrows(IllegalArgumentException.class, () -> {
                excelReader.convertToCsvText(VALID_TEST_INPUT_FILE);
            });
        }

        @Test
        public void negativeSheetIndex() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                ExcelReader.builder().sheetIndex(-5).build();
            });
            assertEquals("SheetIndex cannot be negative", exception.getMessage());
        }

        @Test
        public void settingNullQuoteMode() {
            assertThrows(IllegalArgumentException.class, () -> {
                ExcelReader.builder().quoteMode(null).build();
            });
        }

        @Test
        public void testInvalidSheetName() {
            ExcelReader excelReader = ExcelReader.builder().sheetName("FAKE_WORKSHEET_NAME").build();
            assertThrows(IllegalArgumentException.class, () -> {
                excelReader.convertToCsvText(VALID_TEST_INPUT_FILE);
            });
        }
    }

    @Nested
    @DisplayName("Invalid Password Tests")
    @TestInstance(Lifecycle.PER_CLASS)
    class InvalidPasswordTests {
        @ParameterizedTest
        @CsvSource({
                "bad_password, Password incorrect",
                ",no password was supplied",  // first param is null password
                "'',no password was supplied" // first param is empty string password
        })
        public void invalidPasswordCheck(String password, String expectedMessageSubstring) {
            ExcelReader excelReader = ExcelReader.builder().password(password).build();
            Exception exception = assertThrows(Exception.class, () -> {
                excelReader.convertToCsvText(VALID_TEST_INPUT_PSWD_FILE);
            });
            assertContains(expectedMessageSubstring, exception.getMessage());
        }
    }

    @Test
    public void testSaveCsvOutputFileIllegalNullCharInPath() {
        File outFile = new File("aa().?aa_||._\u0000_bbb.csv");
        ExcelReader excelReader = ExcelReader.builder().build();

        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            excelReader.convertToCsvFile(VALID_TEST_INPUT_FILE, outFile);
        });
        assertContains(
                "Nul character not allowed",
                exception.getMessage());
    }

    private void assertContains(String subString, String mainString) {
        assertTrue(mainString.contains(subString),
                String.format("Expected to find substring '%s' in string '%s'.", subString, mainString));
    }
}
