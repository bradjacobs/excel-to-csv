/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;


public class SheetContent {

    private final String sheetName;

    private final String[][] dataMatrix;

    public SheetContent(String[][] dataMatrix) {
        this("", dataMatrix);
    }

    public SheetContent(String sheetName, String[][] dataMatrix) {
        this.sheetName = sheetName;
        this.dataMatrix = dataMatrix;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String[][] getMatrix() {
        return dataMatrix;
    }

    public List<List<String>> getRows() {
        return Arrays.stream(dataMatrix)
                .map(Arrays::asList)
                .collect(Collectors.toList());
    }
}
