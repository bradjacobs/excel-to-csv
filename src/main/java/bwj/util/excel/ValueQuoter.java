/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package bwj.util.excel;


import org.apache.commons.lang3.StringUtils;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

/**
 * Escape a given value that is compatible for output to a CSV file.
 */
//  NOTE: could have just utilized 'com.fasterxml.jackson.dataformat.csv.CsvMapper'
//    for the quoting logic, but didn't want to have to pull in yet another dependency.
//      (reserve the right to change decision on this)
class ValueQuoter
{
    // any string that has a character below this ascii value will be quoted in 'Normal Mode'
    private static final int NORMAL_CRITERIA_MINIMUM = 45;

    // if a cell value contains any of the characters then the result should be "quoted"
    // before writing to CSV file.  (Lenient Mode)
    private static final Set<Character> MINIMAL_QUOTE_CHARACTERS =
            new HashSet<>(Arrays.asList('"', ',', '\t', '\r', '\n'));

    private final QuoteMode quoteMode;

    public ValueQuoter(QuoteMode quoteMode) {
        this.quoteMode = quoteMode;
    }

    public String applyCsvQuoting(String value) {
        boolean shouldWrap = shouldWrap(value);
        if (shouldWrap) {
            if (value.contains("\"")) {
                value = StringUtils.replace(value,"\"", "\"\"");  // escape embedded double quotes
            }

            // concat this way __only__ b/c could be inside tight loops
            value = new StringBuilder(value.length()+2).append('"').append(value).append('"').toString();
        }
        return value;
    }

    private boolean shouldWrap(String value) {
        if (value == null || value.isEmpty()) {
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
}
