/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.FieldSource;

import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.DASHES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.SPACES;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

public class SpecialCharacterSanitizerTest {
    private static final List<Arguments> spaceChars = Arrays.asList(
            arguments(named("NON_BREAKING SPACE", "\u00a0")),
            arguments(named("EN SPACE", "\u2002")),
            arguments(named("EM SPACE", "\u2003")),
            arguments(named("THREE-PER-EM SPACE", "\u2004")),
            arguments(named("FOUR-PER-EM SPACE", "\u2005")),
            arguments(named("SIX-PER-EM SPACE", "\u2006")),
            arguments(named("FIGURE SPACE", "\u2007")),
            arguments(named("PUNCTUATION SPACE", "\u2008")),
            arguments(named("THIN SPACE", "\u2009")),
            arguments(named("HAIR SPACE", "\u200a")),
            arguments(named("ZERO-WIDTH SPACE", "\u200b")),
            arguments(named("BRAILLE SPACE", "\u2800"))
    );

    // Test that a string with a "special" space character
    //   becomes a "normal" space character
    @ParameterizedTest
    @FieldSource("spaceChars")
    public void sanitizeSpacialSpaceCharacter(String spaceChar) {
        String inputString = "a" + spaceChar + "b";
        String expectedResult = "a b";
        String result = new SpecialCharacterSanitizer(SPACES).sanitize(inputString);
        assertEquals(expectedResult, result, "mismatch result of whitespace char substitution");
    }

    @Test
    public void sanitizeWithDoubleCurlyQuotes() {
        String inputCurlyDoubleQuotes = "she said “hi” to my dog";
        String expectedResult = "she said \"hi\" to my dog";
        String result = new SpecialCharacterSanitizer(QUOTES).sanitize(inputCurlyDoubleQuotes);
        assertEquals(expectedResult, result, "mismatch result of quote character replacement");
    }

    @Test
    public void sanitizeWithSingleCurlyQuotes() {
        String inputCurlySingleQuotes = "she said ‘hi’ to my dog";
        String expectedResult = "she said 'hi' to my dog";
        String result = new SpecialCharacterSanitizer(QUOTES).sanitize(inputCurlySingleQuotes);
        assertEquals(expectedResult, result, "mismatch result of quote character replacement");
    }

    @Test
    public void validateDisablingWhitespaceSanitization() {
        String inputString = "has \u00a0 special space";
        String result = new SpecialCharacterSanitizer(new HashSet<>()).sanitize(inputString);
        assertEquals(inputString, result, "expected input string to remain unchanged.");
    }

    @Test
    public void validateDisablingQuoteSanitization() {
        String inputString = "has special \u201C quote";
        String result = new SpecialCharacterSanitizer(new HashSet<>()).sanitize(inputString);
        assertEquals(inputString, result, "expected input string to remain unchanged.");
    }

    @Test
    public void sanitizeDashCharaters() {
        String inputCurlySingleQuotes = "aaa–bbb－ccc";
        String expectedResult = "aaa-bbb-ccc";
        String result = new SpecialCharacterSanitizer(DASHES).sanitize(inputCurlySingleQuotes);
        assertEquals(expectedResult, result, "mismatch result of quote character replacement");
    }

    @ParameterizedTest
    @CsvSource({
            "lēad, lead",
            "naïve, naive",
            "Façade, Facade",
            "CAFÉ, CAFE",
            "résumé, resume",
            "déjà vu, deja vu"
    })
    void sanitizeBasicDiacritics(String input, String expected) {
        String result = new SpecialCharacterSanitizer(BASIC_DIACRITICS).sanitize(input);
        assertEquals(expected, result, "mismatch expected sanitized diacritics");
    }

    @Nested
    @DisplayName("Default Flag Tests")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class DefaultSanitizeFlagTests {
        // whitespace sanitization ON by default
        @Test
        public void whitespaceDefaultSanitize() {
            String input = "good\u00a0day";
            String expected = "good day";
            String result = new SpecialCharacterSanitizer().sanitize(input);
            assertEquals(expected, result, "mismatch expected sanitized string");
        }

        // quote sanitization ON by default
        @Test
        public void quoteDefaultSanitize() {
            String input = "with “doubles” and ‘singles’";
            String expected = "with \"doubles\" and 'singles'";
            String result = new SpecialCharacterSanitizer().sanitize(input);
            assertEquals(expected, result, "mismatch expected sanitized string");
        }

        // dash sanitization OFF by default
        @Test
        public void dashDefaultSanitize() {
            String input = "Foo–Bar";
            String result = new SpecialCharacterSanitizer().sanitize(input);
            assertEquals(input, result, "mismatch expected sanitized string");
        }

        // diacritics sanitization OFF by default
        @Test
        public void diacriticsDefaultSanitize() {
            String input = "résumé";
            String result = new SpecialCharacterSanitizer().sanitize(input);
            assertEquals(input, result, "mismatch expected sanitized string");
        }
    }

    @Test
    public void validateNullParameter() {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            new SpecialCharacterSanitizer((Set<CharSanitizeFlag>) null);
        });
        assertEquals("Must provide non-null charSanitizeFlags.", exception.getMessage());

        Exception exception2 = assertThrows(IllegalArgumentException.class, () -> {
            new SpecialCharacterSanitizer((CharSanitizeFlag[]) null);
        });
        assertEquals("Must provide non-null charSanitizeFlags.", exception2.getMessage());
    }
}