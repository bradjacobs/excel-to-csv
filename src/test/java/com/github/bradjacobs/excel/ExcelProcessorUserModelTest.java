/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.engine.AbstractExcelReaderTest;


class ExcelProcessorUserModelTest extends AbstractExcelReaderTest<ExcelProcessor, ExcelProcessor.Builder> {

    @Override
    protected ExcelProcessor.Builder createBuilder() {
        return ExcelProcessor.builder().useEventReader(false);
    }

}
