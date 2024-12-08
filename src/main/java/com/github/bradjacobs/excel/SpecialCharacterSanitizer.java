/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import java.text.Normalizer;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import static java.util.stream.Collectors.toMap;

/**
 * Used to convert some special unicode characters into basic equivalent.
 * example: converts 'smart quotes' into normal quotes  (e.g. “” --> "")
 *   and/or convert special space characters (i.e. NBSP characters) to normal spaces.
 */
public class SpecialCharacterSanitizer {

    public enum CharSanitizeFlags {
        SPACES,
        QUOTES,
        BASIC_DIACRITICS
    };

    private static final Character SPACE_CHAR = ' ';
    private static final Set<CharSanitizeFlags> DEFAULT_FLAGS =
            Set.of(CharSanitizeFlags.SPACES, CharSanitizeFlags.QUOTES);

    private final Map<Character,Character> replacementMap;

    public SpecialCharacterSanitizer() {
        this(DEFAULT_FLAGS);
    }

    public SpecialCharacterSanitizer(CharSanitizeFlags... charSanitizeFlags) {
        this(convertToSet(charSanitizeFlags));
    }

    public SpecialCharacterSanitizer(Set<CharSanitizeFlags> charSanitizeFlags) {
        if (charSanitizeFlags == null) {
            throw new IllegalArgumentException("Must provide non-null charSanitizeFlags.");
        }
        this.replacementMap = new HashMap<>();
        if (charSanitizeFlags.contains(CharSanitizeFlags.SPACES)) {
            this.replacementMap.putAll(SPACE_ONLY_REPLACEMENT_MAP);
        }
        if (charSanitizeFlags.contains(CharSanitizeFlags.QUOTES)) {
            this.replacementMap.putAll(QUOTE_ONLY_REPLACEMENT_MAP);
        }
        if (charSanitizeFlags.contains(CharSanitizeFlags.BASIC_DIACRITICS)) {
            this.replacementMap.putAll(DIACRITICS_CHAR_REPLACEMENT_MAP);
        }
    }

    /**
     * TODO - fix description
     * Replace any "special/extended" space characters with the basic space character 0x20.
     * Additionally replace any 'curly quotes' with normal quote (both single and double)
     * NOTE that all 'normal' whitespace characters ['\r', '\n', '\t', ' '] will remain as-is
     * @param input string to sanitize
     * @return sanitized version of the input string
     */
    public String sanitize(String input) {
        if (this.replacementMap.isEmpty()) {
            return input;
        }
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

    private static Set<CharSanitizeFlags> convertToSet(CharSanitizeFlags... charSanitizeFlags) {
        if (charSanitizeFlags == null) {
            return null;
        }
        return new HashSet<>(Arrays.asList(charSanitizeFlags));
    }

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
            '\uFF07', // Misc
    };

    private static final Character[] DOUBLE_QUOTE_CHARS = {
            '\u201C', // "Smart" Double Curved Quote - Left
            '\u201D', // "Smart" Double Curved Quote - Right
            '\u201E', // Low Double Curved Quote - Left
            '\u201F', // Double High-Reversed
            '\u00AB', // Double Guillemet Angle Quote - Left
            '\u00BB', // Double Guillemet Angle Quote - Right
            '\u275D', // Heavy Double Turned Comma Quotation Mark (ornament)
            '\u275E', // Heavy Single Comma Quotation Mark (ornament)
            '\u2826', // Braille Double Closing Quotation Mark
            '\u2834', // Braille Double Opening Quotation Mark
            '\uFF02', // Misc
    };


    private static final Map<Character,Character> SPACE_ONLY_REPLACEMENT_MAP;
    private static final Map<Character,Character> QUOTE_ONLY_REPLACEMENT_MAP;
    private static final Map<Character,Character> DIACRITICS_CHAR_REPLACEMENT_MAP;

    static {
        SPACE_ONLY_REPLACEMENT_MAP = generateSpaceReplacementMap();
        QUOTE_ONLY_REPLACEMENT_MAP = generateQuoteReplacementMap();
        DIACRITICS_CHAR_REPLACEMENT_MAP = generateBasicDiacriticsCharReplacementMap();
    }

    private static Map<Character,Character> generateSpaceReplacementMap() {
        Map<Character, Character> replacementMap = new LinkedHashMap<>();
        for (char c = 0; c < Character.MAX_VALUE; c++) {
            if (Character.isSpaceChar(c) && c != SPACE_CHAR) {
                replacementMap.put(c, SPACE_CHAR);
            }
        }
        replacementMap.put('\u200b', SPACE_CHAR); // ZERO-WIDTH SPACE (tbd if this should be here)
        replacementMap.put('\u2800', SPACE_CHAR); // BRAILLE SPACE
        return replacementMap;
    }

    private static Map<Character,Character> generateQuoteReplacementMap() {
        return new LinkedHashMap<>(){{
            putAll(generateReplacementMap(SINGLE_QUOTE_CHARS, '\''));
            putAll(generateReplacementMap(DOUBLE_QUOTE_CHARS, '"'));
        }};
    }

    private static Map<Character,Character> generateReplacementMap(Character[] inputCharList, Character replacementChar) {
        return Arrays.stream(inputCharList).collect(
                toMap(c -> c, c -> replacementChar));
    }

    /**
     * Creates a lookup replacement map for characters with accents
     * to the 'normal looking' counterpart.
     *   Examples:  'é' -> 'e', 'Ç' -> 'C', 'ö' -> 'o'
     * NOTE: this only considers replacement characters that are in
     *  the basic ascii range < 255
     * @return Map of diacritics character to its replacement value
     */
    private static Map<Character,Character> generateBasicDiacriticsCharReplacementMap() {
        Map<Character,Character> replacementMap = new LinkedHashMap<>();
        Pattern pattern = Pattern.compile("\\p{M}");

        for (char c = 0; c < Character.MAX_VALUE; c++) {
            String input = Character.toString(c);
            String normalized = Normalizer.normalize(input, Normalizer.Form.NFD);
            String output = pattern.matcher(normalized).replaceAll("");

            if (! input.equals(output) && output.length() == 1) {
                char outChar = output.charAt(0);
                if (outChar < 255 && outChar != SPACE_CHAR) {
                    replacementMap.put(c, outChar);
                }
            }
        }

        // want to avoid converting a "not equals" to an "equals", etc
        replacementMap.remove('\u2260'); // remove "not equals"
        replacementMap.remove('\u226E'); // remove "not less than"
        replacementMap.remove('\u226F'); // remove "not greater than"

        return replacementMap;
    }
}
