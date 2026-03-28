/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import org.apache.commons.lang3.Strings;
import org.apache.commons.lang3.Validate;

import java.util.function.Predicate;

/**
 * Given a 2-D string array, escape each value to make csv-compatible and
 * combine to be final csv text string.
 */
public class MatrixToCsvTextConverter {
    private static final String NEW_LINE = System.lineSeparator();

    private final Predicate<String> quoteRule;

    /**
     * Constructor
     * @param quoteMode quoteMode type to determine rules for 'quoting' string values
     */
    public MatrixToCsvTextConverter(QuoteMode quoteMode) {
        Validate.isTrue(quoteMode != null, "QuoteMode cannot be null.");
        this.quoteRule = quoteMode.createPredicate();
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
}
