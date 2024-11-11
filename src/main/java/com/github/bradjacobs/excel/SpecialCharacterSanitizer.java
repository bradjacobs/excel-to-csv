/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import java.util.HashMap;
import java.util.Map;

/**
 * Used to convert some special unicode characters into basic equivalent.
 * Namely, converts 'smart quotes' into normal quotes  (e.g. “” --> "")
 *   and convert special space characters (i.e. NBSP characters) to normal spaces.
 */
public class SpecialCharacterSanitizer {
    private final Map<Character,Character> replacementMap;

    public SpecialCharacterSanitizer() {
        this(true, true);
    }

    /**
     * Constructor
     * @param sanitizeWhiteSpace flag to convert unicode space characters
     * @param sanitizeQuotes flag to convert unicode quote characters
     */
    public SpecialCharacterSanitizer(boolean sanitizeWhiteSpace, boolean sanitizeQuotes) {
        if (sanitizeWhiteSpace && sanitizeQuotes) {
            replacementMap = COMBINED_REPLACEMENT_MAP;
        }
        else if (sanitizeWhiteSpace) {
            replacementMap = SPACE_ONLY_REPLACEMENT_MAP;
        }
        else if (sanitizeQuotes) {
            replacementMap = QUOTE_ONLY_REPLACEMENT_MAP;
        }
        else {
            throw new IllegalArgumentException("Must specify the type of characters to sanitize.");
        }
    }

    /**
     * Replace any "special/extended" space characters with the basic space character 0x20.
     * Additionally replace any 'curly quotes' with normal quote (both single and double)
     * NOTE that all our 'normal' whitespace characters ['\r', '\n', '\t', ' '] will remain as-is
     * @param input string to sanitize
     * @return string with whitespace and quote chars replaced (if any were found)
     */
    public String sanitize(String input) {
        int length = input.length();
        StringBuilder sb = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            char inputCharacter = input.charAt(i);
            if (this.replacementMap.containsKey(inputCharacter)) {
                sb.append(this.replacementMap.get(inputCharacter));
            }
            else {
                sb.append(inputCharacter);
            }
        }
        return sb.toString();
    }

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

    // note: intentionally _not_ considering ascent marks as single quotes.
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

    // dev note: initializing via HashMap instead of creating via Map.of() method
    //   because the performance is actually noticeable with really large files.
    private static final Map<Character,Character> SPACE_ONLY_REPLACEMENT_MAP = new HashMap<>();
    private static final Map<Character,Character> QUOTE_ONLY_REPLACEMENT_MAP = new HashMap<>();
    private static final Map<Character,Character> COMBINED_REPLACEMENT_MAP = new HashMap<>();

    static {
        for (Character c : SPECIAL_SPACE_CHARS) {
            SPACE_ONLY_REPLACEMENT_MAP.put(c, ' ');
        }
        for (Character c : SINGLE_QUOTE_CHARS) {
            QUOTE_ONLY_REPLACEMENT_MAP.put(c, '\'');
        }
        for (Character c : DOUBLE_QUOTE_CHARS) {
            QUOTE_ONLY_REPLACEMENT_MAP.put(c, '"');
        }
        COMBINED_REPLACEMENT_MAP.putAll(SPACE_ONLY_REPLACEMENT_MAP);
        COMBINED_REPLACEMENT_MAP.putAll(QUOTE_ONLY_REPLACEMENT_MAP);
    }
}
