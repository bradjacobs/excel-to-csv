/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.util;

import com.github.bradjacobs.excel.request.SheetInfo;

public class TestSheetInfoUtil {

    public static SheetInfo sheet(String name, int index) {
        return new SheetInfo() {
            @Override
            public String getName() {
                return name;
            }

            @Override
            public int getIndex() {
                return index;
            }
        };
    }
}
