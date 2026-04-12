/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import com.github.bradjacobs.excel.standard.StandardExcelSheetReader;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.function.Executable;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static com.github.bradjacobs.excel.csv.QuoteMode.NORMAL;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileObject;
import static com.github.bradjacobs.excel.util.TestResourceUtil.readResourceFileText;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

class CsvWriterTest {
    // the following are special character that should always be quoted
    private static final List<Character> MINIMAL_QUOTE_CHARACTERS
            = List.of('"', ',', '\t', '\r', '\n');

    private static final CsvWriter normalCsvWriter = new CsvWriter();

    private static final String TEST_OUTPUT_FILE_NAME = "test_output.csv";

    @TempDir
    private Path tempDir;

    @Nested
    class ToCsvStringTests {
        @Test
        public void emptyMatrixToEmptyString() {
            assertEquals("", normalCsvWriter.toCsv(null), "Expected empty string");
            assertEquals("", normalCsvWriter.toCsv(new SheetContent(new String[0][0])), "Expected empty string");
            assertEquals("", normalCsvWriter.toCsv(new SheetContent(new String[1][0])), "Expected empty string");
        }

        @Test
        public void simpleMatrixToString() {
            String[][] matrix = {
                    {"dog", "cow"},
                    {"frog", "cat"}
            };
            String expected = "dog,cow" + System.lineSeparator()
                    + "frog,cat";

            SheetContent sheetContent = new SheetContent(matrix);
            String csvResult = normalCsvWriter.toCsv(sheetContent);
            assertEquals(expected, csvResult, "Mismatch expected CSV output");
        }

        @Test
        public void matrixToStringWithQuoting() {
            String[][] matrix = {
                    {"dog", "say \"hi\""},
                    {"frog", "aa,bb"}
            };
            String expected = "dog,\"say \"\"hi\"\"\"" + System.lineSeparator()
                    + "frog,\"aa,bb\"";

            SheetContent sheetContent = new SheetContent(matrix);
            String csvResult = normalCsvWriter.toCsv(sheetContent);
            assertEquals(expected, csvResult, "Mismatch expected CSV output");
        }

        @Test
        public void noQuotesForBlankValue() {
            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";

            SheetContent sheetContent = new SheetContent(matrix);
            String csvResult = normalCsvWriter.toCsv(sheetContent);
            assertEquals(expected, csvResult, "Mismatch expected CSV output");
        }

    }

    @ParameterizedTest
    @MethodSource("quoteTestProvider")
    public void quoteModeChecking(String input, QuoteMode quoteMode, String expectedOutput) {

        CsvWriter csvWriter = new CsvWriter(quoteMode);
        String[][] matrix = {{input}};

        SheetContent sheetContent = new SheetContent(matrix);
        String output = csvWriter.toCsv(sheetContent);
        assertEquals(expectedOutput, output);
    }

    @Nested
    class SavingCsvFileTests {
        @Test
        public void testSavePathObject() throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";
            SheetContent sheetContent = new SheetContent(matrix);

            CsvWriter.write(testOutputFile, sheetContent);
            assertTrue(Files.exists(testOutputFile), "expected csv file was NOT created");

            String outputFileContent = readFileContents(testOutputFile);
            assertEquals(expected, outputFileContent, "mismatch of content of saved csv file");
        }

        @Test
        public void testSaveFileObject() throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";
            SheetContent sheetContent = new SheetContent(matrix);

            normalCsvWriter.writeToFile(testOutputFile.toFile(), sheetContent);
            assertTrue(Files.exists(testOutputFile), "expected csv file was NOT created");

            String outputFileContent = readFileContents(testOutputFile);
            assertEquals(expected, outputFileContent, "mismatch of content of saved csv file");
        }

        // if the output csv file saved does _NOT_ have any Unicode,
        //  then the 'saveUnicodeFileWithBom' flag should have no effect.
        @Test
        public void testBomFlagWithoutUnicode() throws Exception {
            Path testOutputFileOff1 = tempDir.resolve("test_bom_flag_off.csv");
            Path testOutputFileOn2 = tempDir.resolve("test_bom_flag_on.csv");

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";
            SheetContent sheetContent = new SheetContent(matrix);

            CsvWriter bomFlagOffWriter = new CsvWriter(NORMAL, false);
            CsvWriter bomFlagOnWriter = new CsvWriter(NORMAL, true);

            bomFlagOffWriter.writeToFile(testOutputFileOff1, sheetContent);
            bomFlagOnWriter.writeToFile(testOutputFileOn2, sheetContent);

            assertTrue(Files.exists(testOutputFileOff1), "expected csv file was NOT created");
            assertTrue(Files.exists(testOutputFileOn2), "expected csv file was NOT created");

            String fileContentOff1 = readFileContents(testOutputFileOff1);
            String fileContentOn2 = readFileContents(testOutputFileOn2);

            // first check the actual content is expected
            assertEquals(expected, fileContentOff1, "mismatch of content of saved csv file");
            assertEquals(fileContentOff1, fileContentOn2, "expect 2 files to be the same");
        }

        @Test
        public void testBomFlagWithUnicode() throws Exception {
            Path testOutputFileOff1 = tempDir.resolve("test_bom_flag_off.csv");
            Path testOutputFileOn2 = tempDir.resolve("test_bom_flag_on.csv");

            String[][] matrix = {{"total Façade", "", "in the CAFÉ"}};
            String expected = "\"total Façade\",,\"in the CAFÉ\"";

            SheetContent sheetContent = new SheetContent(matrix);

            CsvWriter bomFlagOffWriter = new CsvWriter(NORMAL, false);
            CsvWriter bomFlagOnWriter = new CsvWriter(NORMAL, true);

            bomFlagOffWriter.writeToFile(testOutputFileOff1, sheetContent);
            bomFlagOnWriter.writeToFile(testOutputFileOn2, sheetContent);

            assertTrue(Files.exists(testOutputFileOff1), "expected csv file was NOT created");
            assertTrue(Files.exists(testOutputFileOn2), "expected csv file was NOT created");

            String fileContentOff1 = readFileContents(testOutputFileOff1);
            String fileContentOn2 = readFileContents(testOutputFileOn2);

            // first check the actual content is expected
            assertEquals(expected, fileContentOff1, "mismatch of content of saved csv file");

            char expectedBom = '\uFEFF';
            // the file that had the bom flag turned on should have it as the first character
            assertEquals(expectedBom, fileContentOn2.charAt(0), "Didn't find expected Bom on saved csv file");

            // now if we remove this first bom character, then the 2 strings should be equal
            assertEquals(fileContentOff1, fileContentOn2.substring(1));
        }

        @Test
        public void saveMultipleFiles() throws Exception {
            String[][] data1 = {
                {"aa", "bb"},
                {"cc", "dd"}
            };
            String expected1 = "aa,bb" + System.lineSeparator() + "cc,dd";
            String[][] data2 = {
                {"ee", "ff"},
                {"gg", "hh"}
            };
            String expected2 = "ee,ff" + System.lineSeparator() + "gg,hh";

            SheetContent sheetContent1 = new SheetContent("data1", data1);
            SheetContent sheetContent2 = new SheetContent("data2", data2);
            List<SheetContent> sheetContentList = List.of(sheetContent1, sheetContent2);

            CsvWriter.write(tempDir, sheetContentList);
            Path file1 = tempDir.resolve("data1.csv");
            Path file2 = tempDir.resolve("data2.csv");
            assertTrue(Files.exists(file1), "expected csv file was NOT created");
            assertTrue(Files.exists(file2), "expected csv file was NOT created");

            String fileContent1 = readFileContents(file1);
            String fileContent2 = readFileContents(file2);
            assertEquals(expected1, fileContent1, "mismatch of content of saved csv file");
            assertEquals(expected2, fileContent2, "mismatch of content of saved csv file");
        }
    }

    private static final String DIR_PATH = Paths.get("").toAbsolutePath().toString();
    protected static List<Arguments> invalidOutputPaths() {
        return Arrays.asList(
                arguments(named("Invalid Output file extension", "outfile.exe"), IllegalArgumentException.class, "Illegal outputFile extension 'exe'.  Must be either 'csv', 'txt' or blank"),
                arguments(named("Invalid Output directory", "/fakeDirectory/myOutputFile.csv"), IllegalArgumentException.class, "Attempted to save CSV output file in a non-existent directory: /fakeDirectory/myOutputFile.csv"),
                arguments(named("Null Csv Output Param", null), IllegalArgumentException.class, "Must supply outputFile location to save CSV data."),
                arguments(named("Directory Output Param", DIR_PATH), IllegalArgumentException.class, "The outputFile cannot be an existing directory.")
        );
    }

    @Nested
    @DisplayName("Exception Behavior")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class ExceptionBehaviorTests {

        @Test
        public void invalidNullQuoteMode() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                new CsvWriter(null);
            });
            assertEquals("QuoteMode cannot be null.", exception.getMessage());
        }

        // todo
        //   non-directory to directory save
        //   directory to directory save
        //   directory to file save
        //   file to directory save
        //   file to file save

        @Test
        public void missingSheetContentParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                normalCsvWriter.writeToDirectory(Path.of("."), null);
            });
            assertEquals("Must supply at least one sheetContent to write.", exception.getMessage());
        }

        @Test
        public void emptySheetContentListParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                normalCsvWriter.writeToDirectory(Path.of("."), List.of());
            });
            assertEquals("Must supply at least one sheetContent to write.", exception.getMessage());
        }

        @ParameterizedTest
        @MethodSource("com.github.bradjacobs.excel.csv.CsvWriterTest#invalidOutputPaths")
        public void invalidOutputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            Path path = location != null ? Paths.get(location) : null;

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            SheetContent sheetContent = new SheetContent(matrix);

            Executable executable1 = () -> normalCsvWriter.writeToFile(path, sheetContent);
            assertExecutableException(executable1, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("com.github.bradjacobs.excel.csv.CsvWriterTest#invalidOutputPaths")
        public void invalidOutputFileParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) throws MalformedURLException {
            File file = location != null ? new File(location) : null;

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            SheetContent sheetContent = new SheetContent(matrix);

            Executable executable1 = () -> normalCsvWriter.writeToFile(file, sheetContent);
            assertExecutableException(executable1, expectedException, expectedMessage);
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


    private static Stream<Arguments> quoteVariations() {
        return Stream.of(
                Arguments.of(QuoteMode.NORMAL, "expected_normal.csv"),
                Arguments.of(QuoteMode.ALWAYS, "expected_always.csv"),
                Arguments.of(QuoteMode.MINIMAL, "expected_minimal.csv"),
                Arguments.of(QuoteMode.NEVER, "expected_never.csv")
        );
    }
    // TODO - this is residual from the older unittest.
    //   CsvWriter tests really shouldn't need use of the SheetReader
    @ParameterizedTest
    @MethodSource("quoteVariations")
    void testExpectedQuoteTextFileParam(QuoteMode quoteMode, String expectedResultFileName) throws Exception {
        String expectedCsvText = readResourceFileText(expectedResultFileName);
        File inputFile = getResourceFileObject("test_data.xlsx");
        ExcelSheetReadRequest request = ExcelSheetReadRequest.from(inputFile).build();

        StandardExcelSheetReader sheetReader = StandardExcelSheetReader.builder().build();
        SheetContent sheetContent = sheetReader.readSheet(request);

        CsvWriter csvWriter = new CsvWriter(quoteMode);
        String csvText = csvWriter.toCsv(sheetContent);

        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    //
    // Test Helper code below...
    //

    private static List<Arguments> quoteTestProvider() {
        List<QuoteTestInput> quoteTestInputList = createQuoteTestList();
        return convertToArgumentList(quoteTestInputList);
    }

    private static List<QuoteTestInput> createQuoteTestList() {
        List<QuoteTestInput> quoteTestInputList = new ArrayList<>();

        // test a string with each type of character that must always be quoted
        for (Character minimalQuoteCharacter : MINIMAL_QUOTE_CHARACTERS) {
            String input = "aaa" + minimalQuoteCharacter + "bbb";
            QuoteTestInput testQuoteInput = createExpectedQuotedTestInput(input);
            quoteTestInputList.add(testQuoteInput);
        }

        // test a string _solely_ with character that must always be quoted
        for (Character minimalQuoteCharacter : MINIMAL_QUOTE_CHARACTERS) {
            String input = Character.toString(minimalQuoteCharacter);
            QuoteTestInput testQuoteInput = createExpectedQuotedTestInput(input);
            quoteTestInputList.add(testQuoteInput);
        }

        // for simple string with spaces, expect it should _not_ be quoted in 'minimal mode'
        String withSpaces = "string with spaces";
        QuoteTestInput testQuoteInput = createExpectedQuotedTestInput(withSpaces);
        testQuoteInput.expectedValueMap.put(QuoteMode.MINIMAL, withSpaces);
        quoteTestInputList.add(testQuoteInput);

        return quoteTestInputList;
    }

    private static List<Arguments> convertToArgumentList(List<QuoteTestInput> quoteTestInputList) {
        List<Arguments> argumentsList = new ArrayList<>();
        for (QuoteTestInput quoteTestInput : quoteTestInputList) {
            quoteTestInput.expectedValueMap.forEach(
                    (k, v) -> argumentsList.add(Arguments.of(quoteTestInput.inputValue, k, v)));
        }
        return argumentsList;
    }

    /**
     * Create a QuoteTestInput where the expected should be same regardless of QuoteMode
     * @param input the input string
     * @return QuoteTestInput object
     */
    private static QuoteTestInput createExpectedQuotedTestInput(String input) {
        String quotedInput = quoteWrap(input);
        Map<QuoteMode,String> expectedMap = new LinkedHashMap<>();
        for (QuoteMode quoteMode : QuoteMode.values()) {
            if (quoteMode.equals(QuoteMode.NEVER)) {
                expectedMap.put(quoteMode, input);
            }
            else {
                expectedMap.put(quoteMode, quotedInput);
            }
        }
        return new QuoteTestInput(input, expectedMap);
    }

    /**
     * Create an expected CSV quoted version of the value
     * @param input input value
     * @return quoted string
     */
    private static String quoteWrap(String input) {
        if (input.contains("\"")) {
            input = input.replace("\"", "\"\"");
        }
        return "\"" + input + "\"";
    }

    private static class QuoteTestInput {
        final String inputValue;
        final Map<QuoteMode,String> expectedValueMap;

        public QuoteTestInput(String inputValue, Map<QuoteMode, String> expectedValueMap) {
            this.inputValue = inputValue;
            this.expectedValueMap = expectedValueMap;
        }
    }

    private String readFileContents(Path filePath) throws IOException {
        return Files.readString(filePath, StandardCharsets.UTF_8);
    }
}