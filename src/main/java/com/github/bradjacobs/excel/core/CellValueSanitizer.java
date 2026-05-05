/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.core;

import com.github.bradjacobs.excel.config.SanitizeType;
import org.apache.commons.lang3.StringUtils;

import java.util.Set;

public class CellValueSanitizer {

    private final boolean trimStringValues;
    private final SpecialCharacterSanitizer specialCharSanitizer;

    public CellValueSanitizer(boolean trimStringValues, Set<SanitizeType> sanitizeTypes) {
        this.trimStringValues = trimStringValues;
        this.specialCharSanitizer = new SpecialCharacterSanitizer(sanitizeTypes);
    }

    public String sanitizeCellValue(String inputValue) {
        if (StringUtils.isEmpty(inputValue)) {
            return "";
        }
        // if there are any certain special Unicode characters (like nbsp or smart quotes),
        // replace w/ normal character equivalent
        String resultValue = specialCharSanitizer.sanitize(inputValue);

        if (this.trimStringValues) {
            resultValue = resultValue.trim();
        }
        return resultValue;
    }
}
