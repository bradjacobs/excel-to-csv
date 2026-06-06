/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import com.github.bradjacobs.excel.request.SheetInfo;

import java.io.InputStream;

public class EventSheet implements SheetInfo {
    private final int index;
    private final String name;
    private final InputStream inputStream;

    public EventSheet(int index, String name, InputStream inputStream) {
        this.index = index;
        this.name = name;
        this.inputStream = inputStream;
    }

    @Override
    public String getName() {
        return name;
    }

    @Override
    public int getIndex() {
        return index;
    }

    public InputStream getInputStream() {
        return inputStream;
    }
}
