/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import org.apache.commons.lang3.Validate;

import java.util.Collection;
import java.util.Locale;

public class ByNameSheetSelector extends ByCollectionSheetSelector<String> {

    public ByNameSheetSelector(String... values) {
        this(toCollection(values));
    }

    public ByNameSheetSelector(Collection<String> values) {
        this(values, "Names");
    }

    protected ByNameSheetSelector(Collection<String> values, String label) {
        super(values, label);
    }

    @Override
    protected String extractSheetInfoKey(SheetInfo sheetInfo) {
        return sheetInfo.getName();
    }

    @Override
    protected String normalizeValue(String value) {
        return value.toLowerCase(Locale.ROOT);
    }

    @Override
    protected void validateCollection(Collection<String> values, String label) {
        super.validateCollection(values, label);
        Validate.isTrue(
                values.stream().noneMatch(String::isEmpty),
                label + " cannot contain empty values");
    }
}
