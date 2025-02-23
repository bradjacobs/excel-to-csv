/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.commons.lang3.StringUtils;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

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

    private final QuoteMode quoteMode;

    /**
     * Constructor
     * @param quoteMode quoteMode type to determine rules for 'quoting' string values
     */
    public MatrixToCsvTextConverter(QuoteMode quoteMode) {
        if (quoteMode == null) {
            throw new IllegalArgumentException("QuoteMode cannot be null.");
        }
        this.quoteMode = quoteMode;
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

                if (shouldQuoteWrap(cellValue)) {
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
     * Determine if the value should be quoted based on the quoteMode
     * @param value value
     * @return true if value should be quoted for csv text.
     */
    private boolean shouldQuoteWrap(String value) {
        if (this.quoteMode.equals(QuoteMode.NEVER)) {
            return false;
        }
        else if (value == null || value.isEmpty()) {
            // going to ignore all blanks for now
            return false;
        }
        else if (this.quoteMode.equals(QuoteMode.ALWAYS)) {
            return true;
        }

        int valueLength = value.length();
        if (this.quoteMode.equals(QuoteMode.NORMAL)) {
            for (int i = 0; i < valueLength; i++) {
                if (value.charAt(i) < NORMAL_CRITERIA_MINIMUM) {
                    return true;
                }
            }
        }
        else if (this.quoteMode.equals(QuoteMode.LENIENT)) {
            for (int i = 0; i < valueLength; i++) {
                if (MINIMAL_QUOTE_CHARACTERS.contains(value.charAt(i))) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * checks if data matrix array is empty
     * @param dataMatrix dataMatrix
     * @return true if dataMatrix is considered 'empty'
     */
    private boolean isEmptyDataMatrix(String[][] dataMatrix) {
        return (dataMatrix == null || dataMatrix.length == 0 || dataMatrix[0].length == 0);
    }
}
