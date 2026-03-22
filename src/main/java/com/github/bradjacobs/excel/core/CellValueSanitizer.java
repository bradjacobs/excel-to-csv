/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.config.SanitizeType;

import java.util.Set;

public class CellValueSanitizer {

    private final boolean autoTrim;
    private final SpecialCharacterSanitizer specialCharSanitizer;

    public CellValueSanitizer(boolean autoTrim, Set<SanitizeType> sanitizeTypes) {
        this.autoTrim = autoTrim;
        this.specialCharSanitizer = new SpecialCharacterSanitizer(sanitizeTypes);
    }

    public String sanitizeCellValue(String inputValue) {
        if (inputValue == null) {
            return "";
        }
        // if there are any certain special Unicode characters (like nbsp or smart quotes),
        // replace w/ normal character equivalent
        String resultValue = specialCharSanitizer.sanitize(inputValue);

        if (this.autoTrim) {
            resultValue = resultValue.trim();
        }
        return resultValue;
    }


}
