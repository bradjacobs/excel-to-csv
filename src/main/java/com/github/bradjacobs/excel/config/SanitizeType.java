/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.config;

/**
 * Enums to define the types of sanitizations
 * to perform on Sheet cell values.
 */
public enum SanitizeType {
    SPACES, // replace special space characters with normal space 0x20 (note '\n','\r','\t' are _NOT_ considered)
    QUOTES, // replace special quotes (like smart quotes) with normal single/double quote character
    DASHES, // replace special dashes with basic dash '-' character
    BASIC_DIACRITICS // replace basic diacritics (e.g. 'é' -> 'e').  Not every possibility is considered
}
