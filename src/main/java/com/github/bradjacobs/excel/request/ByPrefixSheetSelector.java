/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

public class ByPrefixSheetSelector extends ByNameSheetSelector {

    public ByPrefixSheetSelector(String... values) {
        this(toCollection(values));
    }

    public ByPrefixSheetSelector(Collection<String> values) {
        super(values, "Prefixes");
    }

    @Override
    public <S extends SheetInfo> List<S> filterSheets(Map<String, S> sheetMap) {
        List<S> resultList = new ArrayList<>();
        for (String value : valueList) {
            String normalizedValue = normalizeValue(value);
            for (Map.Entry<String, S> entry : sheetMap.entrySet()) {
                if (entry.getKey().startsWith(normalizedValue)
                        && !resultList.contains(entry.getValue())) {
                    resultList.add(entry.getValue());
                }
            }
        }
        return resultList;

        // todo - note this method could technically be implemented like below.
        //   but would need persuasion b/c imho it looks worse.
        // return valueList.stream()
        //         .map(this::normalizeValue)
        //         .flatMap(normalizedValue -> sheetMap.entrySet().stream()
        //                 .filter(entry -> entry.getKey().startsWith(normalizedValue))
        //                 .map(Map.Entry::getValue))
        //         .distinct()
        //         .collect(Collectors.toList());
    }
}
