/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import java.text.Normalizer;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.DASHES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.SPACES;
import static java.util.stream.Collectors.toMap;

/**
 * Used to convert some special unicode characters into basic equivalent.
 * examples: convert 'smart quotes' into normal quotes  (e.g. “” --> "")
 *   or convert special space characters (i.e. NBSP characters) to normal spaces.
 */
public class SpecialCharacterSanitizer {

    // enums to define the types of sanitizations to perform
    public enum CharSanitizeFlag {
        SPACES, // replace special space characters with normal space 0x20 (note '\n','\r','\t' are _NOT_ considered)
        QUOTES, // replace special quotes (like smart quotes) with normal single/double quote character
        DASHES, // replace special dashes with basic dash '-' character
        BASIC_DIACRITICS // replace basic diacritics (e.g. 'é' -> 'e').  Not every possibility is considered
    }

    // by default, handle spaces and quotes
    private static final Set<CharSanitizeFlag> DEFAULT_FLAGS = Set.of(SPACES, QUOTES);
    private static final Character SPACE_CHAR = ' ';

    private final Map<Character,Character> replacementMap;

    public SpecialCharacterSanitizer() {
        this(DEFAULT_FLAGS);
    }

    public SpecialCharacterSanitizer(CharSanitizeFlag... charSanitizeFlags) {
        this(charSanitizeFlags != null ? Arrays.asList(charSanitizeFlags) : null);
    }

    public SpecialCharacterSanitizer(Collection<CharSanitizeFlag> charSanitizeFlags) {
        if (charSanitizeFlags == null) {
            throw new IllegalArgumentException("Must provide non-null charSanitizeFlags.");
        }

        this.replacementMap = new HashMap<>();
        for (CharSanitizeFlag charSanitizeFlag : charSanitizeFlags) {
            this.replacementMap.putAll(FLAG_REPLACEMENT_LOOKUP_MAP.get(charSanitizeFlag));
        }
    }

    /**
     * Replaces 'special characters' with the basic ascii counterpart
     * Works with spaces, quotes, and/or diacritic characters (i.e. 'é' -> 'e')
     * @param input string to sanitize
     * @return sanitized version of the input string
     */
    public String sanitize(String input) {
        if (input.isEmpty() || this.replacementMap.isEmpty()) {
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

    //
    // Below are private static methods for the initial setup of the character lookup maps
    //

    // note: intentionally _not_ considering ascent marks as single quotes.
    private static final Character[] SINGLE_QUOTE_CHARS = {
            '\u02BC', // Letter Apostrophe
            '\u2018', // Single Curved Quote - Left
            '\u2019', // Single Curved Quote - Right
            '\u201A', // Low Single Curved Quote - Left
            '\u201B', // Single High-Reversed
            '\u2032', // Single Prime
            '\u2035', // Reversed Single Prime
            '\u2039', // Single Angle Quote (Guillemet) - Left
            '\u203A', // Single Angle Quote (Guillemet) - Right
            '\u2358', // Apl Functional Symbol Quote Underbar
            '\u235E', // Apl Functional Symbol Quote Quad
            '\u275B', // Heavy Single Turned Comma Quotation Mark (ornament)
            '\u275C', // Heavy Single Comma Quotation Mark (ornament)
            '\u276E', // Heavy Left-Pointing Angle Quotation Mark (ornament)
            '\u276F', // Heavy Right-Pointing Angle Quotation Mark (ornament)
            '\uFF07', // Fullwidth Apostrophe
    };

    private static final Character[] DOUBLE_QUOTE_CHARS = {
            '\u0022', // Ditto
            '\u00AB', // Double Angle Quote (Guillemet) - Left
            '\u00BB', // Double Angle Quote (Guillemet) - Right
            '\u201C', // "Smart" Double Curved Quote - Left
            '\u201D', // "Smart" Double Curved Quote - Right
            '\u201E', // Low Double Curved Quote - Left
            '\u201F', // Double High-Reversed
            '\u2033', // Double Prime
            '\u2036', // Reversed Double Prime
            '\u275D', // Heavy Double Turned Comma Quotation Mark (ornament)
            '\u275E', // Heavy Double Comma Quotation Mark (ornament)
            '\u2826', // Braille Double Closing Quotation Mark
            '\u2834', // Braille Double Opening Quotation Mark
            '\u3003', // Asian Ditto
            '\u301D', // Reversed Double Prime Quotation Mark
            '\u301E', // Double Prime Quotation Mark
            '\u301F', // Low Double Prime Quotation Mark
            '\uFF02', // Fullwidth Quotation Mark
    };

    private static final Character[] DASH_CHARS = {
            '\u05A8', // armenian hyphen
            '\u05BE', // hebrew punctuation maqaf
            '\u1806', // mongolian soft hyphen
            '\u2010', // hyphen
            '\u2011', // non-breaking hyphen
            '\u2012', // figure dash
            '\u2013', // en dash
            '\u2014', // em dash
            '\u2015', // horizontal bar
            '\u207B', // superscript minus
            '\u208B', // subscript minus
            '\u2E3A', // two-em dash
            '\u2E3B', // three-em dash
            '\uFE58', // small em dash
            '\uFE63', // small hyphen-minus
            '\uFF0D', // fullwidth hyphen-minus
    };

    private static final
        Map<CharSanitizeFlag, Map<Character,Character>> FLAG_REPLACEMENT_LOOKUP_MAP
            = Map.of(
                    SPACES, generateSpaceReplacementMap(),
                    QUOTES, generateQuoteReplacementMap(),
                    DASHES, generateDashReplacementMap(),
                    BASIC_DIACRITICS, generateBasicDiacriticsCharReplacementMap()
    );

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

    private static Map<Character,Character> generateDashReplacementMap() {
        return generateReplacementMap(DASH_CHARS, '-');
    }

    private static Map<Character,Character> generateReplacementMap(Character[] inputCharList, Character replacementChar) {
        return Arrays.stream(inputCharList).collect(
                toMap(c -> c, c -> replacementChar));
    }

    /**
     * Creates a lookup replacement map for characters with accents
     * to the 'normal looking' counterpart.
     *   Examples:  'é' -> 'e', 'Ç' -> 'C', 'ö' -> 'o'
     * NOTE1: this only considers replacement characters that are in
     *  the basic ascii range < 255
     * NOTE2: this does not replace most characters that have 'hooks' or 'slashes'
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

        // avoid converting a "not equals" to an "equals", etc
        replacementMap.remove('\u2260'); // remove "not equals"
        replacementMap.remove('\u226E'); // remove "not less than"
        replacementMap.remove('\u226F'); // remove "not greater than"

        return replacementMap;
    }
}