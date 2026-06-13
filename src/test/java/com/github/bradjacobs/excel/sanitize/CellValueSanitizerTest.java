/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.sanitize;

import org.junit.jupiter.api.Test;

import java.util.Set;

import static com.github.bradjacobs.excel.sanitize.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.QUOTES;
import static org.junit.jupiter.api.Assertions.assertEquals;

class CellValueSanitizerTest {

    private static final String ASSERTION_MESSAGE = "mismatch expected cell value string";

    @Test
    void sanitizeCellValueReturnsEmptyStringForNullInput() {
        CellValueSanitizer sanitizer = new CellValueSanitizer(true, Set.of(QUOTES));
        String result = sanitizer.sanitizeCellValue(null);
        assertEquals("", result, ASSERTION_MESSAGE);
    }

    @Test
    void sanitizeCellValueTrimsInputWhenTrimEnabled() {
        CellValueSanitizer sanitizer = new CellValueSanitizer(true, Set.of(QUOTES));
        String inputString = "  dog  ";
        String expectedString = inputString.trim();
        String result = sanitizer.sanitizeCellValue(inputString);
        assertEquals(expectedString, result, ASSERTION_MESSAGE);
    }

    @Test
    void sanitizeCellValuePreservesWhitespaceWhenTrimDisabled() {
        CellValueSanitizer sanitizer = new CellValueSanitizer(false, Set.of(QUOTES));
        String inputString = "  dog  ";
        String result = sanitizer.sanitizeCellValue(inputString);
        assertEquals(inputString, result, ASSERTION_MESSAGE);
    }

    @Test
    void sanitizeCellValueRemovesBasicDiacriticsWhenConfigured() {
        CellValueSanitizer sanitizer = new CellValueSanitizer(true, Set.of(BASIC_DIACRITICS));
        String inputString = "Façade";
        String result = sanitizer.sanitizeCellValue(inputString);
        assertEquals("Facade", result, ASSERTION_MESSAGE);
    }
}
