/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.lang3.Strings;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class MatrixToCsvTextConverterTest {

    // the following are special character that should always be quoted
    private static final List<Character> MINIMAL_QUOTE_CHARACTERS
            = List.of('"', ',', '\t', '\r', '\n');

    private static final MatrixToCsvTextConverter normalConverter
            = new MatrixToCsvTextConverter(QuoteMode.NORMAL);

    @Test
    public void emptyMatrixToEmptyString() {
        assertEquals("", normalConverter.createCsvText(null), "Expected empty string");
        assertEquals("", normalConverter.createCsvText(new String[0][0]), "Expected empty string");
        assertEquals("", normalConverter.createCsvText(new String[1][0]), "Expected empty string");
    }

    @Test
    public void simpleMatrixToString() {
        String[][] matrix = {
                {"dog", "cow"},
                {"frog", "cat"}
        };
        String expected = "dog,cow" + System.lineSeparator()
                + "frog,cat";

        String csvResult = normalConverter.createCsvText(matrix);
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

        String csvResult = normalConverter.createCsvText(matrix);
        assertEquals(expected, csvResult, "Mismatch expected CSV output");
    }

    @Test
    public void noQuotesForBlankValue() {
        String[][] matrix = {{"cow bell", "", "hot dog"}};
        String expected = "\"cow bell\",,\"hot dog\"";

        String csvResult = normalConverter.createCsvText(matrix);
        assertEquals(expected, csvResult, "Mismatch expected CSV output");
    }

    @ParameterizedTest
    @MethodSource("quoteTestProvider")
    public void quoteModeChecking(String input, QuoteMode quoteMode, String expectedOutput) {
        MatrixToCsvTextConverter converter = new MatrixToCsvTextConverter(quoteMode);
        String[][] matrix = {{input}};

        String output = converter.createCsvText(matrix);
        assertEquals(expectedOutput, output);
    }

    @Test
    public void invalidNullQuoteMode() {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            MatrixToCsvTextConverter converter = new MatrixToCsvTextConverter(null);
        });
        assertEquals("QuoteMode cannot be null.", exception.getMessage());
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

        // for simple string with spaces, expect it should _not_ be quoted in 'lenient mode'
        String withSpaces = "string with spaces";
        QuoteTestInput testQuoteInput = createExpectedQuotedTestInput(withSpaces);
        testQuoteInput.expectedValueMap.put(QuoteMode.LENIENT, withSpaces);
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
            input = Strings.CS.replace(input, "\"", "\"\"");
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
}