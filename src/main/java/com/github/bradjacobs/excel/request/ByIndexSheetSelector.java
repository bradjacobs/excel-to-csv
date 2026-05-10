/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import org.apache.commons.lang3.Validate;

import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;

public class ByIndexSheetSelector extends ByCollectionSheetSelector<Integer> {

    public ByIndexSheetSelector(int... values) {
        this(toIntCollection(values));
    }

    public ByIndexSheetSelector(Collection<Integer> values) {
        super(values, "Indexes");
    }

    @Override
    protected Integer extractSheetInfoKey(SheetInfo sheetInfo) {
        return sheetInfo.getIndex();
    }

    @Override
    protected void validateCollection(Collection<Integer> values, String label) {
        super.validateCollection(values, label);
        Validate.isTrue(values.stream().allMatch(n -> n >= 0),
                label + " cannot contain negative values");
    }

    private static List<Integer> toIntCollection(int... values) {
        return values == null ? null : Arrays.stream(values)
                .boxed()
                .collect(Collectors.toList());
    }
}
