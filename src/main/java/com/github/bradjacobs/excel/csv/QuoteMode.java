/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

public enum QuoteMode {
    ALWAYS,  // always quote values (expect for empty/blank)
    NORMAL,  // Quotes if discovers potentially 'unsafe' characters (but can overquote)
    MINIMAL, // only quote if a string contains a character that needs quotes for CSV compliance
    NEVER    // never quote values
}
