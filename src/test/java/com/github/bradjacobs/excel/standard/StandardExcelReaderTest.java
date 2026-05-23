/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.core.AbstractExcelReaderTest;

import java.nio.file.Path;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;

class StandardExcelReaderTest extends AbstractExcelReaderTest<StandardExcelReader, StandardExcelReader.Builder> {

    @Override
    protected StandardExcelReader.Builder createBuilder() {
        return StandardExcelReader.builder();
    }

    private static final Path VALID_TEST_INPUT_PSWD_PATH = getResourceFilePath("test_data_w_pswd_1234.xlsx");
}
