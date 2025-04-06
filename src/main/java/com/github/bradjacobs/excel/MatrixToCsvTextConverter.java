/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.lang3.StringUtils;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.function.Predicate;

import static com.github.bradjacobs.excel.QuoteMode.ALWAYS;
import static com.github.bradjacobs.excel.QuoteMode.LENIENT;
import static com.github.bradjacobs.excel.QuoteMode.NEVER;
import static com.github.bradjacobs.excel.QuoteMode.NORMAL;

/**
 * Given a 2-D string array, escape each value to make csv-compatible and
 * combine to be final csv text string.
 */
public class MatrixToCsvTextConverter {
    // a string that has a character below this ascii value should be quoted.  (Normal Mode)
    private static final int NORMAL_CRITERIA_MINIMUM = 45;

    // a string with any of these explicit characters should be quoted.  (Lenient Mode)
    private static final Set<Character> MINIMAL_QUOTE_CHARACTERS =
            new HashSet<>(Arrays.asList('"', ',', '\t', '\r', '\n'));

    private static final String NEW_LINE = System.lineSeparator();

    private final Predicate<String> shouldQuoteRule;

    /**
     * Constructor
     * @param quoteMode quoteMode type to determine rules for 'quoting' string values
     */
    public MatrixToCsvTextConverter(QuoteMode quoteMode) {
        if (quoteMode == null) {
            throw new IllegalArgumentException("QuoteMode cannot be null.");
        }
        this.shouldQuoteRule = QUOTE_RULE_MAP.get(quoteMode);
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
                if (shouldQuoteRule.test(cellValue)) {
                    // must first escape double quotes
                    if (cellValue.contains("\"")) {
                        cellValue = StringUtils.replace(cellValue,"\"", "\"\"");
                    }
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

    /**
     * checks if data matrix array is empty
     * @param dataMatrix dataMatrix
     * @return true if dataMatrix is considered 'empty'
     */
    private boolean isEmptyDataMatrix(String[][] dataMatrix) {
        return (dataMatrix == null || dataMatrix.length == 0 || dataMatrix[0].length == 0);
    }

    // Below is the logic if a value should be quoted, depending on the quoteMode.
    private static final Predicate<Character> IS_LOW_ASCII_CHAR = c -> c < NORMAL_CRITERIA_MINIMUM;
    private static final Predicate<Character> IS_MIN_CHAR = MINIMAL_QUOTE_CHARACTERS::contains;

    private static final Predicate<String> NEVER_QUOTE_RULE = s -> false;
    private static final Predicate<String> ALWAYS_QUOTE_RULE = s -> !StringUtils.isEmpty(s);
    private static final Predicate<String> NORMAL_QUOTE_RULE = new CharRulePredicate(IS_LOW_ASCII_CHAR);
    private static final Predicate<String> MIN_QUOTE_RULE = new CharRulePredicate(IS_MIN_CHAR);

    private static final Map<QuoteMode, Predicate<String>> QUOTE_RULE_MAP = Map.of(
            NEVER, NEVER_QUOTE_RULE,
            ALWAYS, ALWAYS_QUOTE_RULE,
            NORMAL, NORMAL_QUOTE_RULE,
            LENIENT, MIN_QUOTE_RULE
    );

    private static class CharRulePredicate implements Predicate<String> {
        private final Predicate<Character> charPredicate;
        public CharRulePredicate(Predicate<Character> charPredicate) {
            this.charPredicate = charPredicate;
        }

        @Override
        public boolean test(String value) {
            if (StringUtils.isEmpty(value)) {
                return false;
            }
            for (int i = 0; i < value.length(); i++) {
                if (charPredicate.test(value.charAt(i))) {
                    return true;
                }
            }
            return false;
        }
    }
}
