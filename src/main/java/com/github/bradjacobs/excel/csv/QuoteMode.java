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
        Predicate<String> createPredicate() {
            return value -> !StringUtils.isEmpty(value);
        }
    },
    // Quotes if discovers potentially 'unsafe' characters (but can overquote)
    NORMAL {
        @Override
        Predicate<String> createPredicate() {
            return anyCharMatch(IS_LOW_ASCII_CHAR);
        }
    },
    // only quote if a string contains a character that needs quotes for CSV compliance
    MINIMAL {
        @Override
        Predicate<String> createPredicate() {
            return anyCharMatch(IS_MINIMAL_QUOTE_CHARACTER);
        }
    },
    // never quote values
    NEVER {
        @Override
        Predicate<String> createPredicate() {
            return value -> false;
        }
    };

    abstract Predicate<String> createPredicate();

    // Threshold from jackson-dataformat-csv CsvEncoder (_cfgMinSafeChar).
    // Characters below this ASCII value are considered potentially unsafe.
    private static final int NORMAL_QUOTE_ASCII_THRESHOLD = 45;

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
}
