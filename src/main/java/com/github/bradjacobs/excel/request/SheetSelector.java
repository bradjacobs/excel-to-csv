/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import java.util.List;

public interface SheetSelector {

    <T extends SheetInfo> List<T> filterSheets(List<T> sheets);
}
