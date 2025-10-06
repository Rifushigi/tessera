package org.tessera.util;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Utility class for validating file paths.
 */
public class FileValidator {

    /**
     * Checks if a file exists and is readable.
     *
     * @param path The path to the file.
     * @return true if the file exists and is readable, false otherwise.
     */
    public static boolean fileExistsAndIsReadable(Path path) {
        return path != null && Files.exists(path) && Files.isReadable(path);
    }

    /**
     * Checks if a path is a valid directory and creates it if it doesn't exist.
     *
     * @param path The path to the directory.
     * @return true if the directory exists or was created successfully, false otherwise.
     */
    public static boolean createDirectoryIfNotExists(Path path) {
        if (path == null) {
            return false;
        }
        if (Files.exists(path)) {
            return Files.isDirectory(path);
        } else {
            try {
                Files.createDirectories(path);
                return true;
            } catch (IOException e) {
                return false;
            }
        }
    }

    /**
     * Checks if a file has the expected extension.
     *
     * @param path      The path to the file.
     * @param extension The required file extension e.g., ".docx".
     * @return true if the file extension matches, false otherwise.
     */
    public static boolean hasExtension(Path path, String extension) {
        if (path == null || extension == null) {
            return false;
        }
        String fileName = path.getFileName().toString();
        return fileName.toLowerCase().endsWith(extension.toLowerCase());
    }
}
