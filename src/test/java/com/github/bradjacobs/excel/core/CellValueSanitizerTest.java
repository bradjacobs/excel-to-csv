/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import org.junit.jupiter.api.Test;

import java.util.Set;

import static com.github.bradjacobs.excel.config.SanitizeType.QUOTES;
import static org.junit.jupiter.api.Assertions.assertEquals;

public class CellValueSanitizerTest {

    // the advanced implementation can pass in a null to
    //   this method, so ensure it's handled correctly.
    @Test
    public void sanitizeCellNullValue() {
        CellValueSanitizer sanitizer = new CellValueSanitizer(true, Set.of(QUOTES));
        String result = sanitizer.sanitizeCellValue(null);
        assertEquals("", result, "mismatch expected cell value");
    }

    // TODO add more tests
}
