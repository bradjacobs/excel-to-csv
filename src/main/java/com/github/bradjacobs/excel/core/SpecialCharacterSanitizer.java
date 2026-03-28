/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.config.SanitizeType;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Validate;

import java.text.Normalizer;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import static com.github.bradjacobs.excel.config.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.config.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.config.SanitizeType.SPACES;
import static java.util.stream.Collectors.toMap;

/**
 * Used to convert some special Unicode characters into basic equivalent.
 * examples: convert 'smart quotes' into normal quotes  (e.g. “” --> "")
 *   or convert special space characters (i.e. NBSP characters) to normal spaces.
 */
public class SpecialCharacterSanitizer {

    // by default, handle spaces and quotes
    public static final Set<SanitizeType> DEFAULT_FLAGS = Set.of(SPACES, QUOTES);
    private static final Character SPACE_CHAR = ' ';

    private final Map<Character, Character> replacementMap;

    public SpecialCharacterSanitizer() {
        this(DEFAULT_FLAGS);
    }

    public SpecialCharacterSanitizer(SanitizeType... sanitizeTypes) {
        this(sanitizeTypes != null ? Arrays.asList(sanitizeTypes) : null);
    }

    public SpecialCharacterSanitizer(Collection<SanitizeType> sanitizeTypes) {
        Validate.isTrue(sanitizeTypes != null, "Must provide non-null sanitizeTypes.");
        this.replacementMap = buildReplacementMap(sanitizeTypes);
    }

    /**
     * Replaces 'special characters' with the basic ascii counterpart
     * Works with spaces, quotes, and/or diacritic characters (i.e. 'é' -> 'e')
     *
     * @param input string to sanitize
     * @return sanitized version of the input string
     */
    public String sanitize(String input) {
        if (StringUtils.isEmpty(input) || replacementMap.isEmpty()) {
            return input;
        }

        int length = input.length();
        StringBuilder sb = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            char inputCharacter = input.charAt(i);
            Character replacementCharacter = replacementMap.get(inputCharacter);
            sb.append(replacementCharacter != null ? replacementCharacter : inputCharacter);
        }
        return sb.toString();
    }

    /**
     * Generate replacement map for the given charSanitizeFlags.
     */
    private static Map<Character, Character> buildReplacementMap(Collection<SanitizeType> sanitizeTypes) {
        Map<Character, Character> map = new HashMap<>();
        for (SanitizeType sanitizeType : sanitizeTypes) {
            map.putAll(FLAG_REPLACEMENT_LOOKUP_MAP.get(sanitizeType));
        }
        return map;
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

    // Only considering characters that look like a 'single horizontal line',
    // such as dashes, hyphens, minus signs and so on.  Does _not_ include
    // 'wavy' characters, tildes, vertical dashes, etc.
    // The criteria for what qualifies as a dash here is obviously VERY subjective.
    private static final Character[] DASH_CHARS = {
            //'\u00AD', // soft hyphen (Note: include or not?  usually invisible)
            '\u02D7', // modifier letter minus sign
            '\u05A8', // Armenian hyphen
            '\u05BE', // Hebrew punctuation maqaf
            '\u1806', // Mongolian soft hyphen
            '\u2010', // hyphen
            '\u2011', // non-breaking hyphen
            '\u2012', // figure dash
            '\u2013', // en dash
            '\u2014', // em dash
            '\u2015', // horizontal bar
            '\u2043', // hyphen bullet
            '\u207B', // superscript minus
            '\u208B', // subscript minus
            '\u2212', // minus sign
            '\u23AF', // horizontal line extension
            '\u23BA', // horizontal scan line-1
            '\u23BB', // horizontal scan line-3
            '\u23BC', // horizontal scan line-7
            '\u23BD', // horizontal scan line-9
            '\u23E4', // straightness
            '\u2500', // box drawing light horizontal line
            '\u2501', // box drawing heavy horizontal line
            '\u268A', // monogram for yang
            '\u2796', // heavy minus sign
            '\u2E3A', // two-em dash
            '\u2E3B', // three-em dash
            '\uFE58', // small em dash
            '\uFE63', // small hyphen-minus
            '\uFF0D', // fullwidth hyphen-minus
    };

    private static final
    Map<SanitizeType, Map<Character, Character>> FLAG_REPLACEMENT_LOOKUP_MAP
            = Map.of(
                    SPACES, generateSpaceReplacementMap(),
                    QUOTES, generateQuoteReplacementMap(),
                    DASHES, generateDashReplacementMap(),
                    BASIC_DIACRITICS, generateBasicDiacriticsReplacementMap()
    );

    private static Map<Character,Character> generateSpaceReplacementMap() {
        Map<Character, Character> replacementMap = new LinkedHashMap<>();
        for (char c = 0; c < Character.MAX_VALUE; c++) {
            if (Character.isSpaceChar(c) && c != SPACE_CHAR) {
                replacementMap.put(c, SPACE_CHAR);
            }
        }
        replacementMap.put('\u2800', SPACE_CHAR); // BRAILLE SPACE
        // NOTE: 'zero width' spaces are NOT counted as spaces
        //   '\u200b' - ZERO-WIDTH SPACE
        //   '\ufeff' - ZERO WIDTH NO-BREAK SPACE
        //   '\u180e' - MONGOLIAN VOWEL SEPARATOR
        return replacementMap;
    }

    private static Map<Character,Character> generateQuoteReplacementMap() {
        Map<Character, Character> resultMap = new LinkedHashMap<>();
        resultMap.putAll(generateReplacementMap(SINGLE_QUOTE_CHARS, '\''));
        resultMap.putAll(generateReplacementMap(DOUBLE_QUOTE_CHARS, '"'));
        return resultMap;
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
     * Examples:  'é' -> 'e', 'Ç' -> 'C', 'ö' -> 'o'
     * NOTE1: this only considers replacement characters that are in
     * the basic/extended ascii range < 255
     * NOTE2: this does not replace most characters that have 'hooks' or 'slashes'
     *
     * @return Map of diacritics character to its replacement value
     */
    private static Map<Character, Character> generateBasicDiacriticsReplacementMap() {
        Map<Character, Character> replacementMap = new LinkedHashMap<>();
        Pattern pattern = Pattern.compile("\\p{M}");

        for (char c = 0; c < Character.MAX_VALUE; c++) {
            String input = Character.toString(c);
            String normalized = Normalizer.normalize(input, Normalizer.Form.NFD);
            String output = pattern.matcher(normalized).replaceAll("");

            if (!input.equals(output) && output.length() == 1) {
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