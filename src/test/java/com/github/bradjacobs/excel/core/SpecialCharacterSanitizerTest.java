/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.config.SanitizeType;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.function.Executable;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.FieldSource;

import java.util.Arrays;
import java.util.List;
import java.util.Set;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.config.SanitizeType.SPACES;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

class SpecialCharacterSanitizerTest {
    private static final String SANITIZED_STRING_MISMATCH = "mismatch expected sanitized string";
    private static final String UNCHANGED_STRING_EXPECTED = "expected input string to remain unchanged.";
    private static final String NON_NULL_SANITIZE_TYPES_REQUIRED = "Must provide non-null sanitizeTypes.";
    private static final Set<SanitizeType> NO_SANITIZATION_TYPES = Set.of();

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
            arguments(named("BRAILLE SPACE", "\u2800"))
    );

    // Test that a string with a "special" space character
    //   becomes a "normal" space character
    @ParameterizedTest
    @FieldSource("spaceChars")
    public void sanitizeSpecialSpaceCharacter(String spaceChar) {
        assertSanitizedEquals("a b", "a" + spaceChar + "b", SPACES);
    }

    @Test
    public void sanitizeWithDoubleCurlyQuotes() {
        assertSanitizedEquals("she said \"hi\" to my dog", "she said “hi” to my dog", QUOTES);
    }

    @Test
    public void sanitizeWithSingleCurlyQuotes() {
        assertSanitizedEquals("she said 'hi' to my dog", "she said ‘hi’ to my dog", QUOTES);
    }

    @Test
    public void validateDisablingWhitespaceSanitization() {
        assertUnchanged("has \u00a0 special space", NO_SANITIZATION_TYPES);
    }

    @Test
    public void validateDisablingQuoteSanitization() {
        assertUnchanged("has special \u201C quote", NO_SANITIZATION_TYPES);
    }

    @Test
    public void sanitizeDashCharacters() {
        String inputWithDashCharacters = "aaa–bbb－ccc˗d−e";
        assertSanitizedEquals("aaa-bbb-ccc-d-e", inputWithDashCharacters, DASHES);
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
    public void sanitizeBasicDiacritics(String input, String expected) {
        assertSanitizedEquals(expected, input, BASIC_DIACRITICS);
    }

    @Test
    public void unsupportedSanitizeType() {
        // a 'null' is the only way to test invalid param
        //   error message for an unsupported enum type
        Exception exception = assertThrows(IllegalArgumentException.class, () ->
                new SpecialCharacterSanitizer(Arrays.asList(DASHES, null))
        );
        assertEquals("No replacement map registered for SanitizeType: null", exception.getMessage());
    }

    @Nested
    @DisplayName("Default Flag Tests")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class DefaultSanitizeFlagTests {
        // whitespace sanitization ON by default
        @Test
        public void whitespaceDefaultSanitize() {
            assertDefaultSanitizedEquals("good day", "good\u00a0day");
        }

        // quote sanitization ON by default
        @Test
        public void quoteDefaultSanitize() {
            assertDefaultSanitizedEquals("with \"doubles\" and 'singles'", "with “doubles” and ‘singles’");
        }

        // dash sanitization OFF by default
        @Test
        public void dashDefaultSanitize() {
            assertDefaultSanitizedEquals("Foo–Bar", "Foo–Bar");
        }

        // diacritics sanitization OFF by default
        @Test
        public void diacriticsDefaultSanitize() {
            assertDefaultSanitizedEquals("résumé", "résumé");
        }
    }

    @Test
    public void validateNullTypesParameter() {
        assertInvalidSanitizeTypes(() -> new SpecialCharacterSanitizer((Set<SanitizeType>) null));
        assertInvalidSanitizeTypes(() -> new SpecialCharacterSanitizer((SanitizeType[]) null));
    }

    @Test
    public void sanitizeNullString() {
        // normal usage won't try to sanitize a 'null', but test for completeness.
        assertNull(sanitize(null, SPACES), "expected a 'null' to be sanitized to a 'null'");
    }

    private static String sanitize(String input, SanitizeType... sanitizeTypes) {
        return new SpecialCharacterSanitizer(sanitizeTypes).sanitize(input);
    }

    private static String sanitize(String input, Set<SanitizeType> sanitizeTypes) {
        return new SpecialCharacterSanitizer(sanitizeTypes).sanitize(input);
    }

    private static void assertSanitizedEquals(String expected, String input, SanitizeType... sanitizeTypes) {
        assertEquals(expected, sanitize(input, sanitizeTypes), SANITIZED_STRING_MISMATCH);
    }

    private static void assertUnchanged(String input, Set<SanitizeType> sanitizeTypes) {
        assertEquals(input, sanitize(input, sanitizeTypes), UNCHANGED_STRING_EXPECTED);
    }

    private static void assertDefaultSanitizedEquals(String expected, String input) {
        assertEquals(expected, new SpecialCharacterSanitizer().sanitize(input), SANITIZED_STRING_MISMATCH);
    }

    private static void assertInvalidSanitizeTypes(Executable constructorCall) {
        Exception exception = assertThrows(IllegalArgumentException.class, constructorCall);
        assertEquals(NON_NULL_SANITIZE_TYPES_REQUIRED, exception.getMessage());
    }
}