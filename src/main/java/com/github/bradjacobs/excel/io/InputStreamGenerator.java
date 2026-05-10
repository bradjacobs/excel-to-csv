/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.io;

import org.apache.commons.lang3.Validate;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Set;
import java.util.zip.GZIPInputStream;

public class InputStreamGenerator {
    private static final Set<String> VALID_URL_SCHEMES =
            Set.of("http", "https", "ftp", "file");
    private static final int CONNECTION_TIMEOUT = 20000;
    // some websites require a userAgent value set.
    private static final String USER_AGENT_VALUE = "javaClient/" + System.getProperty("java.version");

    public InputStream getInputStream(File inputFile) throws IOException {
        Validate.isTrue(inputFile != null, "Must provide an input file.");
        return getInputStream(inputFile.toPath());
    }

    public InputStream getInputStream(Path inputFile) throws IOException {
        Validate.isTrue(inputFile != null, "Must provide an input file.");
        Validate.isTrue(!Files.isDirectory(inputFile), "The input file cannot be a directory.");
        if (!Files.exists(inputFile)) {
            // throw a different exception for FileNotFound
            throw new FileNotFoundException(String.format("Invalid Excel file path: %s", inputFile.toAbsolutePath()));
        }
        return new BufferedInputStream( Files.newInputStream(inputFile) );
    }

    public InputStream getInputStream(URL url) throws IOException {
        Validate.isTrue(url != null, "Must provide an input url.");

        String urlProtocol = url.getProtocol();
        Validate.isTrue(VALID_URL_SCHEMES.contains(urlProtocol), "URL has an unsupported protocol: %s", urlProtocol);

        if (urlProtocol.equalsIgnoreCase("file")) {
            return getInputStream( Paths.get(url.getPath()) );
        }

        // Could switch to an httpClient in the future.
        URLConnection connection = openConnection(url);
        connection.setConnectTimeout(CONNECTION_TIMEOUT);
        connection.setReadTimeout(CONNECTION_TIMEOUT);
        connection.setRequestProperty("User-Agent", USER_AGENT_VALUE);
        connection.setRequestProperty("Accept-Encoding", "gzip");
        String encoding = connection.getContentEncoding();
        InputStream urlInputStream = connection.getInputStream();
        if ("gzip".equalsIgnoreCase(encoding)) {
            urlInputStream = new GZIPInputStream(urlInputStream);
        }
        return new BufferedInputStream(urlInputStream);
    }

    // scoped to allow for mock testing
    protected URLConnection openConnection(URL url) throws IOException {
        return url.openConnection();
    }
}
