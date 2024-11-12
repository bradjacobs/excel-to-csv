/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.lang3.StringUtils;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class MatrixToCsvTextConverterTest {

    private static final MatrixToCsvTextConverter normalConverter
            = new MatrixToCsvTextConverter(QuoteMode.NORMAL);

    @Test
    public void testEmptyInput() {
        assertEquals("", normalConverter.createCsvText(null), "Expected empty string");
        assertEquals("", normalConverter.createCsvText(new String[0][0]), "Expected empty string");
        assertEquals("", normalConverter.createCsvText(new String[1][0]), "Expected empty string");
    }

    @Test
    public void testSimpleMatrix() {
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
    public void testSimpleMatrixRequiringQuotes() {
        String[][] matrix = {
                {"dog", "say \"hi\""},
                {"frog", "aa,bb"}
        };
        String expected = "dog,\"say \"\"hi\"\"\"" + System.lineSeparator()
                + "frog,\"aa,bb\"";

        String csvResult = normalConverter.createCsvText(matrix);
        assertEquals(expected, csvResult, "Mismatch expected CSV output");
    }

    @ParameterizedTest
    @MethodSource("quoteTestProvider")
    public void testQuoteRules(String input, QuoteMode quoteMode, String expectedOutput) {
        MatrixToCsvTextConverter converter = new MatrixToCsvTextConverter(quoteMode);
        String[][] matrix = {{input}};

        String output = converter.createCsvText(matrix);
        assertEquals(expectedOutput, output);
    }

    //
    // Test Helper code below...
    //

    private static List<Arguments> quoteTestProvider() {
        List<QuoteTestInput> quoteTestInputList = createQuoteTestList();
        return convertToArgumentList(quoteTestInputList);
    }

    // the following are special character that should always be quoted
    private static final List<Character> MINIMAL_QUOTE_CHARACTERS
            = List.of('"', ',', '\t', '\r', '\n');

    private static List<QuoteTestInput> createQuoteTestList() {
        List<QuoteTestInput> quoteTestInputList = new ArrayList<>();

        // test a string with each type of character that must always be quoted
        for (Character minimalQuoteCharacter : MINIMAL_QUOTE_CHARACTERS) {
            String input = "aaa" + minimalQuoteCharacter + "bbb";
            QuoteTestInput testQuoteInput = createAlwaysQuoteInput(input);
            quoteTestInputList.add(testQuoteInput);
        }
        // test a string _solely_ with character that must always be quoted
        for (Character minimalQuoteCharacter : MINIMAL_QUOTE_CHARACTERS) {
            String input = Character.toString(minimalQuoteCharacter);
            QuoteTestInput testQuoteInput = createAlwaysQuoteInput(input);
            quoteTestInputList.add(testQuoteInput);
        }

        String withSpaces = "string with spaces";
        QuoteTestInput testQuoteInput = createAlwaysQuoteInput(withSpaces);
        testQuoteInput.expectedValueMap.put(QuoteMode.LENIENT, withSpaces);
        quoteTestInputList.add(testQuoteInput);

        return quoteTestInputList;
    }

    private static List<Arguments> convertToArgumentList(List<QuoteTestInput> quoteTestInputList) {
        List<Arguments> argumentsList = new ArrayList<>();
        for (QuoteTestInput quoteTestInput : quoteTestInputList) {
            Map<QuoteMode, String> expectedMap = quoteTestInput.expectedValueMap;
            for (Map.Entry<QuoteMode, String> entry : expectedMap.entrySet()) {
                argumentsList.add(Arguments.of(quoteTestInput.inputValue, entry.getKey(), entry.getValue()));
            }
        }
        return argumentsList;
    }

    /**
     * Create a QuoteTestInput where the expected should be same regardless of QuoteMode
     * @param input the input string
     * @return QuoteTestInput object
     */
    private static QuoteTestInput createAlwaysQuoteInput(String input) {
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
            input = StringUtils.replace(input,"\"", "\"\"");
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