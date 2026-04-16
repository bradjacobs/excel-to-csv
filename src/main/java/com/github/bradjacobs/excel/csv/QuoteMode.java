/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import org.apache.commons.lang3.StringUtils;

import java.util.function.IntPredicate;
import java.util.function.Predicate;

/**
 * Defines different quote modes for generating csv text string.
 */
public enum QuoteMode {
    // always quote values (expect for empty/blank)
    ALWAYS {
        @Override
        Predicate<String> createPredicate(char delimiter) {
            return value -> !StringUtils.isEmpty(value);
        }
    },
    // Quotes if discovers potentially 'unsafe' characters (but can overquote)
    NORMAL {
        @Override
        Predicate<String> createPredicate(char delimiter) {
            return createNormalQuotePredicate(delimiter);
        }
    },
    // only quote if a string contains a character that needs quotes for CSV compliance
    MINIMAL {
        @Override
        Predicate<String> createPredicate(char delimiter) {
            return createMinimalPredicate(delimiter);
        }
    },
    // never quote values
    NEVER {
        @Override
        Predicate<String> createPredicate(char delimiter) {
            return value -> false;
        }
    };

    private static final char DEFAULT_DELIMITER = ',';

    public Predicate<String> createPredicate() {
        return createPredicate(DEFAULT_DELIMITER);
    }

    abstract Predicate<String> createPredicate(char delimiter);

    // Threshold from jackson-dataformat-csv CsvEncoder (_cfgMinSafeChar).
    // Characters below this ASCII value are considered potentially unsafe.
    private static final int NORMAL_QUOTE_ASCII_THRESHOLD = 45;

    /**
     * Simple predicate to check if any character in a string matches.
     * Will return false for null or empty strings.
     * @param characterPredicate character predicate (as IntPredicate object)
     * @return String Predicate
     */
    private static Predicate<String> anyCharMatch(IntPredicate characterPredicate) {
        return s -> s != null && s.chars().anyMatch(characterPredicate);
    }

    private static Predicate<String> createNormalQuotePredicate(char delimiter) {
        IntPredicate charPredicate = delimiter < NORMAL_QUOTE_ASCII_THRESHOLD
                ? c -> c < NORMAL_QUOTE_ASCII_THRESHOLD
                : c -> c < NORMAL_QUOTE_ASCII_THRESHOLD || c == delimiter;
        return anyCharMatch(charPredicate);
    }

    private static Predicate<String> createMinimalPredicate(char delimiter) {
        IntPredicate minQuoteCharPredicate = isMinQuoteCharPredicate(delimiter);
        return anyCharMatch(minQuoteCharPredicate);
    }

    private static IntPredicate isMinQuoteCharPredicate(char delimiter) {
        return c -> c == '"'
                || c == delimiter
                || c == '\t'
                || c == '\r'
                || c == '\n';
    }
}
