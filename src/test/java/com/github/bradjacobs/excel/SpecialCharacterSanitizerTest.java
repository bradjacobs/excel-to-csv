/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.FieldSource;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

public class SpecialCharacterSanitizerTest {
    // NOTE:  (*) means Character.isWhitespace() == false
    private static final List<Arguments> spaceChars = Arrays.asList(
            arguments(named("NON_BREAKING SPACE", "\u00a0")), // (*)
            arguments(named("EN SPACE", "\u2002")),
            arguments(named("EM SPACE", "\u2003")),
            arguments(named("THREE-PER-EM SPACE", "\u2004")),
            arguments(named("FOUR-PER-EM SPACE", "\u2005")),
            arguments(named("SIX-PER-EM SPACE", "\u2006")),
            arguments(named("FIGURE SPACE", "\u2007")), // (*)
            arguments(named("PUNCTUATION SPACE", "\u2008")),
            arguments(named("THIN SPACE", "\u2009")),
            arguments(named("HAIR SPACE", "\u200a")),
            arguments(named("ZERO-WIDTH SPACE", "\u200b")), // (*)
            arguments(named("BRAILLE SPACE", "\u2800")) // (*)
    );

    // Test that a string with a "special" space character
    //   becomes a "normal" space character
    @ParameterizedTest
    @FieldSource("spaceChars")
    public void testConvertToNormalSpace(String spaceChar) {
        String inputString = "a" + spaceChar + "b";
        String expectedResult = "a b";

        String result = new SpecialCharacterSanitizer().sanitize(inputString);
        assertEquals(expectedResult, result, "mismatch result of whitespace char substitution");
    }

    @Test
    public void testSanitizeDoubleCurlyQuotes() {
        String inputCurlyDoubleQuotes = "she said “hi” to my dog";
        String expectedResult = "she said \"hi\" to my dog";

        String result = new SpecialCharacterSanitizer().sanitize(inputCurlyDoubleQuotes);
        assertEquals(expectedResult, result, "mismatch result of quote character replacement");
    }

    @Test
    public void testSanitizeSingleCurlyQuotes() {
        String inputCurlySingleQuotes = "she said ‘hi’ to my dog";
        String expectedResult = "she said 'hi' to my dog";

        String result = new SpecialCharacterSanitizer().sanitize(inputCurlySingleQuotes);
        assertEquals(expectedResult, result, "mismatch result of quote character replacement");
    }

    @Test
    public void testDontSanitizeWhitespaceIfDisabled() {
        String inputString = "has \u00a0 special space";
        String result = new SpecialCharacterSanitizer(false, true).sanitize(inputString);
        assertEquals(inputString, result, "expected input string to remain unchanged.");
    }

    @Test
    public void testDontSanitizeQuotesIfDisabled() {
        String inputString = "has special \u201C quote";
        String result = new SpecialCharacterSanitizer(true, false).sanitize(inputString);
        assertEquals(inputString, result, "expected input string to remain unchanged.");
    }

    @Test
    public void testExceptionWithBadConstructorParams() {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            new SpecialCharacterSanitizer(false, false);
        });
        assertEquals("Must specify the type of characters to sanitize.", exception.getMessage());
    }

    private static final Character[] SINGLE_QUOTE_CHARS = {
        '\u2018', // Single Curved Quote - Left
        '\u2019', // Single Curved Quote - Right
        '\u201A', // Low Single Curved Quote - Left
        '\u201B', // Single High-Reversed
        '\u2039', // Single Guillemet Angle Quote - Left
        '\u203A', // Single Guillemet Angle Quote - Right
        '\u275B', // Heavy Single Turned Comma Quotation Mark (ornament)
        '\u275C', // Heavy Single Comma Quotation Mark (ornament)
    };

    private static final Character[] DOUBLE_QUOTE_CHARS = {
        '\u201C', // "Smart" Double Curved Quote - Left
        '\u201D', // "Smart" Double Curved Quote - Right
        '\u201E', // Low Double Curved Quote - Left
        '\u201E', // Double High-Reversed
        '\u00AB', // Double Guillemet Angle Quote - Left
        '\u00BB', // Double Guillemet Angle Quote - Right
        '\u275D', // Heavy Double Turned Comma Quotation Mark (ornament)
        '\u275E', // Heavy Single Comma Quotation Mark (ornament)
        '\u2826', // Braille Double Closing Quotation Mark
        '\u2834', // Braille Double Opening Quotation Mark
    };
}
