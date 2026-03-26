/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

interface SheetVisibilityPolicy {

    default boolean isRowVisible(int rowNum) {
        return true;
    }

    default boolean isColumnVisible(int columnIndex) {
        return true;
    }

    default boolean isCellVisible(int rowNum, int columnIndex) {
        return isRowVisible(rowNum) && isColumnVisible(columnIndex);
    }
}