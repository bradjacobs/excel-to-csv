package com.github.bradjacobs.excel;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

public class WhitespaceSanitizer {
    // "special" space characters that will be converted
    //   to a "normal" space character.  (*) means Character.isWhitespace() == false
    private static final Character[] SPECIAL_SPACE_CHARS = {
            '\u00a0', // NON_BREAKING SPACE (*),
            '\u2002', // EN SPACE
            '\u2003', // EM SPACE
            '\u2004', // THREE-PER-EM SPACE
            '\u2005', // FOUR-PER-EM SPACE
            '\u2006', // SIX-PER-EM SPACE
            '\u2007', // FIGURE SPACE (*)
            '\u2008', // PUNCTUATION SPACE
            '\u2009', // THIN SPACE
            '\u200a', // HAIR SPACE
            '\u200b', // ZERO-WIDTH SPACE (*)
            '\u2800'  // BRAILLE SPACE (*)
    };

    // note: can fix syntax when upgrade the JDK version
    private static final Set<Character> SPECIAL_SPACE_CHAR_SET = new HashSet<>(Arrays.asList(SPECIAL_SPACE_CHARS));

    /**
     * Replace any "special/extended" space characters with the basic space character 0x20,
     * then also do a normal "trim()"
     * NOTE that all our the 'normal' whitespace characters ['\r', '\n', '\t', ' '] will remain as-is
     * @param input string to sanitize
     * @return string with whitespace chars replaces (if any were found)
     */
    public String sanitize(String input) {
        // TODO: probably not the best way to do this!... but works for now.
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < input.length(); i++) {
            char inputCharacter = input.charAt(i);
            if (SPECIAL_SPACE_CHAR_SET.contains(inputCharacter)) {
                sb.append(' ');
            }
            else {
                sb.append(inputCharacter);
            }
        }
        return sb.toString().trim();
    }
}
