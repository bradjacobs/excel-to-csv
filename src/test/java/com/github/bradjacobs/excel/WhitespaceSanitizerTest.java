/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.FieldSource;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Named.named;
import static org.junit.jupiter.params.provider.Arguments.arguments;

public class WhitespaceSanitizerTest {
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

        String result = new WhitespaceSanitizer().sanitize(inputString);
        assertEquals(expectedResult, result, "mismatch result of whitespace char substitution");
    }

    // Test that a string will get 'trimmed' correctly with any "special" space characters
    @ParameterizedTest
    @FieldSource("spaceChars")
    public void testTrimSpecialSpace(String spaceChar) {
        String inputString = "a" + spaceChar;
        String expectedResult = "a";

        String result = new WhitespaceSanitizer().sanitize(inputString);
        assertEquals(expectedResult, result, "mismatch result of whitespace char trim");
    }
}
