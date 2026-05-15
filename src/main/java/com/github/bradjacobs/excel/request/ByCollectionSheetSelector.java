/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import org.apache.commons.lang3.Validate;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;

import static java.util.function.Function.identity;

abstract public class ByCollectionSheetSelector<T> implements SheetSelector {

    protected final List<T> valueList;

    protected ByCollectionSheetSelector(Collection<T> values, String valueTypeLabel) {
        validateCollection(values, valueTypeLabel);
        this.valueList = new ArrayList<>(values);
    }

    protected abstract T extractSheetInfoKey(SheetInfo sheetInfo);

    private T mapKey(SheetInfo s) {
        return normalizeValue(extractSheetInfoKey(s));
    }

    @Override
    public <S extends SheetInfo> List<S> filterSheets(List<S> sheets) {
        Map<T, S> sheetMap = sheets.stream().collect(
                Collectors.toMap(this::mapKey,
                        identity(),
                        (oldValue, newValue) -> { throw new IllegalStateException("Duplicate key found: " + oldValue); },
                        LinkedHashMap::new));
        return filterSheets(sheetMap);
    }

    public <S extends SheetInfo> List<S> filterSheets(Map<T, S> sheetMap) {
        return valueList.stream()
                .map(value -> Optional.ofNullable(sheetMap.get(normalizeValue(value)))
                        .orElseThrow(() -> new IllegalArgumentException("Requested Excel sheet not found: '" + value + "'")))
                .collect(Collectors.toList());
    }


        // Each subclass can normalize its key if needed
    protected T normalizeValue(T value) {
        return value; // default = no change
    }

    protected void validateCollection(Collection<T> values, String label) {
        Validate.isTrue(values != null, label + " cannot be null");
        Validate.isTrue(!values.isEmpty(), label + " cannot be empty");
        Validate.isTrue(values.stream().noneMatch(Objects::isNull), label + " cannot contain null values");
        Set<T> valueSet = values.stream()
                .map(this::normalizeValue)
                .collect(Collectors.toSet());
        // Currently prefer to tell user of duplicates
        //   rather than automatically remove.
        Validate.isTrue(values.size() == valueSet.size(),
                label + " cannot contain duplicate values");
    }

    @SafeVarargs
    protected static <T> Collection<T> toCollection(T... values) {
        return values == null ? null : Arrays.asList(values);
    }
}
