/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import java.util.List;

public class AllSheetsSelector implements SheetSelector{

    @Override
    public <T extends SheetInfo> List<T> filterSheets(List<T> sheets) {
        return sheets;
    }
}
