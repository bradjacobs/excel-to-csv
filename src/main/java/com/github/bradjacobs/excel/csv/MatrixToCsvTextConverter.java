/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Strings;

import java.util.function.IntPredicate;
import java.util.function.Predicate;

/**
 * Given a 2-D string array, escape each value to make csv-compatible and
 * combine to be final csv text string.
 */
public class MatrixToCsvTextConverter {
    private static final int NORMAL_QUOTE_ASCII_THRESHOLD = 45;
    private static final String NEW_LINE = System.lineSeparator();

    private final Predicate<String> quoteRule;

    /**
     * Constructor
     * @param quoteMode quoteMode type to determine rules for 'quoting' string values
     */
    public MatrixToCsvTextConverter(QuoteMode quoteMode) {
        this.quoteRule = getQuoteRule(quoteMode);
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
     * Simple predicate to check if any character in a string matches.
     * Will return false for null or empty strings.
     * @param characterPredicate character predicate (as IntPredicate object)
     * @return String Predicate
     */
    public static Predicate<String> anyCharMatch(IntPredicate characterPredicate) {
        return s -> s != null && s.chars().anyMatch(characterPredicate);
    }

    private static final IntPredicate IS_LOW_ASCII_CHAR = c -> c < NORMAL_QUOTE_ASCII_THRESHOLD;
    private static final IntPredicate IS_MINIMAL_QUOTE_CHARACTER = c ->
            c == '"' || c == ',' || c == '\t' || c == '\r' || c == '\n';

    private static final Predicate<String> NEVER_QUOTE_RULE = value -> false;
    private static final Predicate<String> ALWAYS_QUOTE_RULE = value -> !StringUtils.isEmpty(value);
    private static final Predicate<String> NORMAL_QUOTE_RULE = anyCharMatch(IS_LOW_ASCII_CHAR);
    private static final Predicate<String> MINIMAL_QUOTE_RULE = anyCharMatch(IS_MINIMAL_QUOTE_CHARACTER);

    /**
     * Lookup the the quote rule predicate for the given quoteMode.
     * @param quoteMode quoteMode
     * @return the quote rule predicate
     */
    private static Predicate<String> getQuoteRule(QuoteMode quoteMode) {
        if (quoteMode == null) {
            throw new IllegalArgumentException("QuoteMode cannot be null.");
        }
        switch (quoteMode) {
            case NEVER:
                return NEVER_QUOTE_RULE;
            case ALWAYS:
                return ALWAYS_QUOTE_RULE;
            case NORMAL:
                return NORMAL_QUOTE_RULE;
            case MINIMAL:
                return MINIMAL_QUOTE_RULE;
            default:
                throw new IllegalArgumentException("Unsupported quoteMode: " + quoteMode);
        }
    }
}
