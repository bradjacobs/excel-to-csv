/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Strings;

import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.function.Predicate;

import static com.github.bradjacobs.excel.csv.QuoteMode.ALWAYS;
import static com.github.bradjacobs.excel.csv.QuoteMode.LENIENT;
import static com.github.bradjacobs.excel.csv.QuoteMode.NEVER;
import static com.github.bradjacobs.excel.csv.QuoteMode.NORMAL;

/**
 * Given a 2-D string array, escape each value to make csv-compatible and
 * combine to be final csv text string.
 */
public class MatrixToCsvTextConverter {
    private static final int NORMAL_QUOTE_ASCII_THRESHOLD = 45;
    private static final Set<Character> LENIENT_QUOTE_CHARACTERS =
            new HashSet<>(Set.of('"', ',', '\t', '\r', '\n'));
    private static final String NEW_LINE = System.lineSeparator();

    private final Predicate<String> quoteRule;

    /**
     * Constructor
     * @param quoteMode quoteMode type to determine rules for 'quoting' string values
     */
    public MatrixToCsvTextConverter(QuoteMode quoteMode) {
        if (quoteMode == null) {
            throw new IllegalArgumentException("QuoteMode cannot be null.");
        }
        this.quoteRule = QUOTE_RULE_MAP.get(quoteMode);
    }

    /**
     * Convert the 2-D value matrix into a single CSV string.
     *   (quotes will be applied on values as needed)
     * @param dataMatrix string data matrix
     * @return csv file string
     */
    public String createCsvText(String[][] dataMatrix) {
        if (isEmptyDataMatrix(dataMatrix)) {
            return "";
        }

        StringBuilder sb = new StringBuilder();
        int columnCount = dataMatrix[0].length;
        int lastColumnIndex = columnCount - 1;

        for (String[] rowData : dataMatrix) {
            if (sb.length() != 0) {
                sb.append(NEW_LINE);
            }
            for (int i = 0; i < columnCount; i++) {
                String cellValue = rowData[i];
                if (quoteRule.test(cellValue)) {
                    // must first escape double quotes
                    cellValue = escapeDoubleQuotes(cellValue);
                    sb.append('\"').append(cellValue).append('\"');
                }
                else {
                    sb.append(cellValue);
                }

                if (i != lastColumnIndex) {
                    sb.append(',');
                }
            }
        }

        return sb.toString();
    }

    private String escapeDoubleQuotes(String value) {
        if (value.contains("\"")) {
            return Strings.CS.replace(value, "\"", "\"\"");
        }
        return value;
    }

    /**
     * checks if data matrix array is empty
     * @param dataMatrix dataMatrix
     * @return true if dataMatrix is considered 'empty'
     */
    private boolean isEmptyDataMatrix(String[][] dataMatrix) {
        return dataMatrix == null || dataMatrix.length == 0 || dataMatrix[0].length == 0;
    }

    /**
     * Check if any characters in the given string match the character predicate.
     * @param value input string
     * @param characterRule Predicate used to test each character in the string.
     * @return true if there is at least one character match.
     */
    private static boolean containsMatchingCharacter(String value, Predicate<Character> characterRule) {
        if (StringUtils.isEmpty(value)) {
            return false;
        }

        for (int i = 0; i < value.length(); i++) {
            if (characterRule.test(value.charAt(i))) {
                return true;
            }
        }
        return false;
    }

    private static final Predicate<Character> IS_LOW_ASCII_CHAR =
            character -> character < NORMAL_QUOTE_ASCII_THRESHOLD;
    private static final Predicate<Character> IS_LENIENT_QUOTE_CHARACTER =
            LENIENT_QUOTE_CHARACTERS::contains;

    private static final Predicate<String> NEVER_QUOTE_RULE = value -> false;
    private static final Predicate<String> ALWAYS_QUOTE_RULE = value -> !StringUtils.isEmpty(value);
    private static final Predicate<String> NORMAL_QUOTE_RULE =
            value -> containsMatchingCharacter(value, IS_LOW_ASCII_CHAR);
    private static final Predicate<String> LENIENT_QUOTE_RULE =
            value -> containsMatchingCharacter(value, IS_LENIENT_QUOTE_CHARACTER);

    private static final Map<QuoteMode, Predicate<String>> QUOTE_RULE_MAP = Map.of(
            NEVER, NEVER_QUOTE_RULE,
            ALWAYS, ALWAYS_QUOTE_RULE,
            NORMAL, NORMAL_QUOTE_RULE,
            LENIENT, LENIENT_QUOTE_RULE
    );
}
