/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import com.github.bradjacobs.excel.api.BasicSheetContent;
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
import org.junit.jupiter.params.provider.ValueSource;

import java.io.File;
import java.io.IOException;
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

    private static final CsvWriter DEFAULT_CSV_WRITER = CsvWriter.builder().build();

    private static final String TEST_OUTPUT_FILE_NAME = "test_output.csv";
    private static final String FILE_NOT_CREATED_MESSAGE = "expected csv file was NOT created";
    private static final String CSV_FILE_CONTENT_MISMATCH_MESSAGE = "mismatch of content of saved csv file";

    @TempDir
    private Path tempDir;

    @Nested
    class ToCsvStringTests {
        @Test
        public void emptyMatrixToEmptyString() {
            assertEquals("", DEFAULT_CSV_WRITER.toCsv(null), "Expected empty string");
            assertEquals("", DEFAULT_CSV_WRITER.toCsv(new BasicSheetContent(new String[0][0])), "Expected empty string");
            assertEquals("", DEFAULT_CSV_WRITER.toCsv(new BasicSheetContent(new String[1][0])), "Expected empty string");
        }

        @Test
        public void simpleMatrixToString() {
            String[][] matrix = {
                    {"dog", "cow"},
                    {"frog", "cat"}
            };
            String expected = "dog,cow" + System.lineSeparator()
                    + "frog,cat";

            String csvResult = DEFAULT_CSV_WRITER.toCsv(sheetContent(matrix));
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

            String csvResult = DEFAULT_CSV_WRITER.toCsv(sheetContent(matrix));
            assertEquals(expected, csvResult, "Mismatch expected CSV output");
        }

        @Test
        public void noQuotesForBlankValue() {
            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";

            String csvResult = DEFAULT_CSV_WRITER.toCsv(sheetContent(matrix));
            assertEquals(expected, csvResult, "Mismatch expected CSV output");
        }

        @Test
        public void convertWithNullValue() {
            // under 'normal circumstances' the matrix should never
            // contain a null, but this is a sanity check to ensure we
            // don't blow up on it.
            String[][] matrix = {{"cow bell", null, "hot dog"}};
            String expected = "\"cow bell\",null,\"hot dog\"";

            String csvResult = DEFAULT_CSV_WRITER.toCsv(sheetContent(matrix));
            assertEquals(expected, csvResult, "Mismatch expected CSV output");
        }

        @ParameterizedTest(name = "Delimiter Quote test Normal Mode with input ''{0}''")
        @ValueSource(strings = {",", "\t", ";", ":", "|"})
        void quoteNormalSpecialDelimiter(String delimiter) {
            char delimiterChar = delimiter.charAt(0);

            String valueWithDelimiter = "aa" + delimiter + "bb";
            String valueWithNormalSafeChars = "abc";
            String valueWithUnsafeCharBelowThreshold = "aa bb";

            String[][] matrix = {
                    {valueWithDelimiter},
                    {valueWithNormalSafeChars},
                    {valueWithUnsafeCharBelowThreshold}
            };

            CsvWriter csvWriter = CsvWriter.builder()
                    .quoteMode(QuoteMode.NORMAL)
                    .delimiter(delimiterChar)
                    .build();
            String output = csvWriter.toCsv(sheetContent(matrix));

            String expected = "\"" + valueWithDelimiter + "\""
                    + System.lineSeparator()
                    + valueWithNormalSafeChars
                    + System.lineSeparator()
                    + "\"" + valueWithUnsafeCharBelowThreshold + "\"";
            assertEquals(expected, output);
        }

        @ParameterizedTest(name = "Delimiter Quote test Minimal Mode with input ''{0}''")
        @ValueSource(strings = {",", "\t", ";", ":", "|"})
        void quoteMinimalSpecialDelimiter(String delimiter) {
            char delimiterChar = delimiter.charAt(0);

            String valueWithDelimiter = "aa" + delimiter + "bb";
            String valueWithNormalSafeChars = "abc";

            String[][] matrix = {
                    {valueWithDelimiter},
                    {valueWithNormalSafeChars}
            };

            CsvWriter csvWriter = CsvWriter.builder()
                    .quoteMode(QuoteMode.MINIMAL)
                    .delimiter(delimiterChar)
                    .build();
            String output = csvWriter.toCsv(sheetContent(matrix));

            String expected = "\"" + valueWithDelimiter + "\""
                    + System.lineSeparator()
                    + valueWithNormalSafeChars;
            assertEquals(expected, output);
        }
    }

    @ParameterizedTest
    @MethodSource("quoteTestProvider")
    public void quoteModeChecking(String input, QuoteMode quoteMode, String expectedOutput) {

        CsvWriter csvWriter = CsvWriter.builder().quoteMode(quoteMode).build();
        String[][] matrix = {{input}};

        String output = csvWriter.toCsv(sheetContent(matrix));
        assertEquals(expectedOutput, output);
    }

    @Nested
    class SavingCsvFileTests {
        @Test
        public void testSavePathObject() throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";

            CsvWriter.writeToFile(testOutputFile, sheetContent(matrix));
            assertFileContent(testOutputFile, expected);
        }

        @Test
        public void testSaveFileObject() throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";

            DEFAULT_CSV_WRITER.saveToFile(testOutputFile.toFile(), sheetContent(matrix));
            assertFileContent(testOutputFile, expected);
        }

        // if the output csv file saved does _NOT_ have any Unicode,
        //  then the 'saveUnicodeFileWithBom' flag should have no effect.
        @Test
        public void testBomFlagWithoutUnicode() throws Exception {
            Path testOutputFileOff1 = tempDir.resolve("test_bom_flag_off.csv");
            Path testOutputFileOn2 = tempDir.resolve("test_bom_flag_on.csv");

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";
            SheetContent sheetContent = sheetContent(matrix);

            CsvWriter bomFlagOffWriter = CsvWriter.builder().saveUnicodeFileWithBom(false).build();
            CsvWriter bomFlagOnWriter = CsvWriter.builder().saveUnicodeFileWithBom(true).build();

            bomFlagOffWriter.saveToFile(testOutputFileOff1, sheetContent);
            bomFlagOnWriter.saveToFile(testOutputFileOn2, sheetContent);

            assertFileContent(testOutputFileOff1, expected);
            assertFileContent(testOutputFileOn2, expected);
        }

        @Test
        public void testBomFlagWithUnicode() throws Exception {
            Path testOutputFileOff1 = tempDir.resolve("test_bom_flag_off.csv");
            Path testOutputFileOn2 = tempDir.resolve("test_bom_flag_on.csv");

            String[][] matrix = {{"total Façade", "", "in the CAFÉ"}};
            String expected = "\"total Façade\",,\"in the CAFÉ\"";
            String expectedWithBom = "\uFEFF" + expected;

            CsvWriter bomFlagOffWriter = CsvWriter.builder().saveUnicodeFileWithBom(false).build();
            CsvWriter bomFlagOnWriter = CsvWriter.builder().saveUnicodeFileWithBom(true).build();

            SheetContent sheetContent = sheetContent(matrix);
            bomFlagOffWriter.saveToFile(testOutputFileOff1, sheetContent);
            bomFlagOnWriter.saveToFile(testOutputFileOn2, sheetContent);

            assertFileContent(testOutputFileOff1, expected);
            assertFileContent(testOutputFileOn2, expectedWithBom);
        }

        @Test
        public void saveMultipleFilesToDirectoryPath() throws Exception {
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

            SheetContent sheetContent1 = new BasicSheetContent("data1", data1);
            SheetContent sheetContent2 = new BasicSheetContent("data2", data2);
            List<SheetContent> sheetContentList = List.of(sheetContent1, sheetContent2);

            CsvWriter.writeToDirectory(tempDir, sheetContentList);

            Path file1 = tempDir.resolve("data1.csv");
            Path file2 = tempDir.resolve("data2.csv");
            assertFileContent(file1, expected1);
            assertFileContent(file2, expected2);
        }

        @Test
        public void saveAllowFileOverwrite() throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            String expected = "\"cow bell\",,\"hot dog\"";

            DEFAULT_CSV_WRITER.saveToFile(testOutputFile.toFile(), sheetContent(matrix));
            assertFileContent(testOutputFile, expected);

            String[][] matrix2 = {{"abc bell", "", "def dog"}};
            String expected2 = "\"abc bell\",,\"def dog\"";

            CsvWriter customWriter = CsvWriter.builder().allowOverwriteFile(true).build();
            customWriter.saveToFile(testOutputFile.toFile(), sheetContent(matrix2));
            assertFileContent(testOutputFile, expected2);
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
        public void dontAllowFileOverwrite() throws Exception {
            Path testOutputFile = tempDir.resolve(TEST_OUTPUT_FILE_NAME);

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            SheetContent sheetContent = sheetContent(matrix);

            DEFAULT_CSV_WRITER.saveToFile(testOutputFile.toFile(), sheetContent);
            assertTrue(Files.exists(testOutputFile), "the expected csv file was NOT created");

            // by default, allow Overwrite file = false
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                DEFAULT_CSV_WRITER.saveToFile(testOutputFile.toFile(), sheetContent);
            });
            assertEquals("Attempted to overwrite an existing file: " + TEST_OUTPUT_FILE_NAME, exception.getMessage());
        }

        @Test
        public void invalidNullQuoteMode() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                CsvWriter.builder().quoteMode(null).build();
            });
            assertEquals("QuoteMode cannot be null.", exception.getMessage());
        }

        @Test
        public void missingSheetContentParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                DEFAULT_CSV_WRITER.saveToDirectory(Path.of("."), null);
            });
            assertEquals("Must supply at least one sheetContent to write.", exception.getMessage());
        }

        @Test
        public void emptySheetContentListParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                DEFAULT_CSV_WRITER.saveToDirectory(Path.of("."), List.of());
            });
            assertEquals("Must supply at least one sheetContent to write.", exception.getMessage());
        }

        @Test
        public void invalidDelimiterParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                CsvWriter.builder().delimiter('a').build();
            });
            assertEquals("Invalid delimiter: a", exception.getMessage());
        }

        @Test
        public void nullOutputFilePathParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                String[][] matrix = {{"cow bell", "", "hot dog"}};
                DEFAULT_CSV_WRITER.saveToFile((Path)null, sheetContent(matrix));
            });
            assertEquals("Must supply outputFile location to save CSV data.", exception.getMessage());
        }

        @Test
        public void nullOutputDirectoryParameter() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                String[][] matrix = {{"cow bell", "", "hot dog"}};
                DEFAULT_CSV_WRITER.saveToDirectory((Path)null, List.of(sheetContent(matrix)));
            });
            assertEquals("Must supply outputDirectory location to save CSV files.", exception.getMessage());
        }

        @Test
        public void nullSheetNameOnMultiSave() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                String[][] matrix = {{"cow bell", "", "hot dog"}};
                SheetContent sheetContent1 = new BasicSheetContent("sheet", matrix);
                SheetContent sheetContent2 = new BasicSheetContent(null, matrix);
                List<SheetContent> sheetContentList = List.of(sheetContent1, sheetContent2);
                DEFAULT_CSV_WRITER.saveToDirectory(Path.of("."), sheetContentList);
            });
            assertEquals("Must supply a non-empty sheetName for each sheetContent to write.", exception.getMessage());
        }
        @Test
        public void emptySheetNameOnMultiSave() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                String[][] matrix = {{"cow bell", "", "hot dog"}};
                SheetContent sheetContent1 = new BasicSheetContent("sheet", matrix);
                SheetContent sheetContent2 = new BasicSheetContent("", matrix);
                List<SheetContent> sheetContentList = List.of(sheetContent1, sheetContent2);
                DEFAULT_CSV_WRITER.saveToDirectory(Path.of("."), sheetContentList);
            });
            assertEquals("Must supply a non-empty sheetName for each sheetContent to write.", exception.getMessage());
        }
        @Test
        public void emptyNullInListOnMultiSave() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                String[][] matrix = {{"cow bell", "", "hot dog"}};
                List<SheetContent> sheetContentList = new ArrayList<>();
                sheetContentList.add(new BasicSheetContent("sheet1", matrix));
                sheetContentList.add(null);
                sheetContentList.add(new BasicSheetContent("sheet2", matrix));
                DEFAULT_CSV_WRITER.saveToDirectory(Path.of("."), sheetContentList);
            });
            assertEquals("Must supply a non-empty sheetName for each sheetContent to write.", exception.getMessage());
        }

        @ParameterizedTest
        @MethodSource("com.github.bradjacobs.excel.csv.CsvWriterTest#invalidOutputPaths")
        public void invalidOutputPathParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            Path path = location != null ? Paths.get(location) : null;

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            SheetContent sheetContent = sheetContent(matrix);

            Executable executable1 = () -> DEFAULT_CSV_WRITER.saveToFile(path, sheetContent);
            assertExecutableException(executable1, expectedException, expectedMessage);
        }

        @ParameterizedTest
        @MethodSource("com.github.bradjacobs.excel.csv.CsvWriterTest#invalidOutputPaths")
        public void invalidOutputFileParameter(String location, Class<? extends Exception> expectedException, String expectedMessage) {
            File file = location != null ? new File(location) : null;

            String[][] matrix = {{"cow bell", "", "hot dog"}};
            SheetContent sheetContent = sheetContent(matrix);

            Executable executable1 = () -> DEFAULT_CSV_WRITER.saveToFile(file, sheetContent);
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

        CsvWriter csvWriter = CsvWriter.builder().quoteMode(quoteMode).build();
        String csvText = csvWriter.toCsv(sheetContent);

        assertEquals(expectedCsvText, csvText, "mismatch of expected csv output");
    }

    //
    // Test Helper code below...
    //

    private static SheetContent sheetContent(String[][] matrix) {
        return new BasicSheetContent(matrix);
    }

    private void assertFileExists(Path filePath) {
        assertTrue(Files.exists(filePath), FILE_NOT_CREATED_MESSAGE);
    }

    private void assertFileContent(Path filePath, String expectedContent) throws IOException {
        assertFileExists(filePath);
        assertEquals(expectedContent, readFileContents(filePath), CSV_FILE_CONTENT_MISMATCH_MESSAGE);
    }


    private static List<Arguments> quoteTestProvider() {
        List<QuoteTestInput> quoteTestInputList = createQuoteTestList();
        return toArguments(quoteTestInputList);
    }

    private static List<QuoteTestInput> createQuoteTestList() {
        List<QuoteTestInput> quoteTestInputList = new ArrayList<>();

        // test a string with each type of character that must always be quoted
        for (Character minimalQuoteCharacter : MINIMAL_QUOTE_CHARACTERS) {
            String input = "aaa" + minimalQuoteCharacter + "bbb";
            QuoteTestInput testQuoteInput = createExpectedQuotedTestInput(input);
            quoteTestInputList.add(testQuoteInput);
        }

        // test a string _solely_ with a character that must always be quoted
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

    private static List<Arguments> toArguments(List<QuoteTestInput> quoteTestInputList) {
        List<Arguments> argumentsList = new ArrayList<>();
        for (QuoteTestInput quoteTestInput : quoteTestInputList) {
            quoteTestInput.expectedValueMap.forEach(
                    (k, v) -> argumentsList.add(Arguments.of(quoteTestInput.inputValue, k, v)));
        }
        return argumentsList;
    }

    /**
     * Create a QuoteTestInput where the expected should be the same regardless of QuoteMode
     * @param input the input string
     * @return QuoteTestInput object
     */
    private static QuoteTestInput createExpectedQuotedTestInput(String input) {
        String quotedInput = quoteCsvValue(input);
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
     * Create an expected CSV-quoted version of the value
     * @param input input value
     * @return quoted string
     */
    private static String quoteCsvValue(String input) {
        if (input.contains("\"")) {
            input = input.replace("\"", "\"\"");
        }
        return "\"" + input + "\"";
    }

    private static class QuoteTestInput {
        final String inputValue;
        final Map<QuoteMode, String> expectedValueMap;

        public QuoteTestInput(String inputValue, Map<QuoteMode, String> expectedValueMap) {
            this.inputValue = inputValue;
            this.expectedValueMap = expectedValueMap;
        }
    }

    private String readFileContents(Path filePath) throws IOException {
        return Files.readString(filePath, StandardCharsets.UTF_8);
    }
}