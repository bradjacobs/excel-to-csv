package com.github.bradjacobs.excel.util;

import org.apache.commons.io.IOUtils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 * Helper class to read in files from the test resources directory.
 */
public class TestResourceUtil {

    public static URL getResourceFileUrl(String fileName) {
        URL resourceUrl = TestResourceUtil.class.getClassLoader().getResource(fileName);
        assertNotNull(resourceUrl, "Unable to locate file resource: " + fileName);
        return resourceUrl;
    }

    public static Path getResourceFilePath(String fileName) {
        return Paths.get(URI.create(getResourceFileUrl(fileName).toString()));
    }

    public static File getResourceFileObject(String fileName) {
        return new File( getResourceFileUrl(fileName).getPath() );
    }

    /**
     * Read in the text body for the given resource file
     * @param fileName fileName
     * @return the file content as a string
     */
    public static String readResourceFileText(String fileName) {
        try (InputStream is = TestResourceUtil.class.getClassLoader().getResourceAsStream(fileName)) {
            assertNotNull(is, "Unable to read file: " + fileName);
            return IOUtils.toString(is, StandardCharsets.UTF_8);
        }
        catch (IOException e) {
            throw new RuntimeException(String.format("Unable to read test resource file: %s.  Reason: %s", fileName, e.getMessage()), e);
        }
    }
}
