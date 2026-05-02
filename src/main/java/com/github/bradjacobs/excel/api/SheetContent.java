/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.util.List;

public interface SheetContent {

    String getSheetName();

    int getRowCount();

    int getColumnCount();

    boolean isEmpty();

    String getCellValue(int rowIndex, int columnIndex);

    List<String> getRow(int rowIndex);

    List<List<String>> getRows();

    String[][] getMatrix();
}