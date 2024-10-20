/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import static org.testng.Assert.assertEquals;

public class WhitespaceSanitizerTest {
    // NOTE:  (*) means Character.isWhitespace() == false
    @DataProvider(name = "space_chars")
    public Object[][] spaceCharDataParams(){
        return new Object[][] {
                {"\u00a0"},// NON_BREAKING SPACE (*)
                {"\u2002"}, //EN SPACE
                {"\u2003"}, //EM SPACE
                {"\u2004"}, // THREE-PER-EM SPACE
                {"\u2005"}, // FOUR-PER-EM SPACE
                {"\u2006"}, // SIX-PER-EM SPACE
                {"\u2007"}, // FIGURE SPACE (*)
                {"\u2008"}, // PUNCTUATION SPACE
                {"\u2009"}, // THIN SPACE
                {"\u200a"}, // HAIR SPACE
                {"\u200b"}, // ZERO-WIDTH SPACE (*)
                {"\u2800"}, // BRAILLE SPACE (*)
        };
    }

    // Test that a string with a "special" space character
    //   becomes a "normal" space character
    @Test(dataProvider = "space_chars")
    public void testConvertToNormalSpace(String spaceChar) throws Exception {
        String inputString = "a" + spaceChar + "b";
        String expectedResult = "a b";

        String result = new WhitespaceSanitizer().sanitize(inputString);
        assertEquals(result, expectedResult, "mismatch result of whitespace char substitution");
    }

    // Test that a string will get 'trimmed' correctly with any "special" space characters
    //
    @Test(dataProvider = "space_chars")
    public void testTrimSpecialSpace(String spaceChar) throws Exception {
        String inputString = "a" + spaceChar;
        String expectedResult = "a";

        String result = new WhitespaceSanitizer().sanitize(inputString);
        assertEquals(result, expectedResult, "mismatch result of whitespace char trim");
    }
}
