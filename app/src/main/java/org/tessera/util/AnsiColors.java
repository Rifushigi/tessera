package org.tessera.util;


/**
 * Utility class for adding ANSI color codes to console output.
 */
public class AnsiColors {

    public static final String RESET = "\u001B[0m";
    public static final String GREEN = "\u001B[32m";
    public static final String YELLOW = "\u001B[33m";
    public static final String RED = "\u001B[31m";
    public static final String CYAN = "\u001B[36m";

    /**
     * Applies a color to the given text.
     *
     * @param color The ANSI color code.
     * @param text  The text to color.
     * @return The colored text string.
     */
    public static String colored(String color, String text) {
        return color + text + RESET;
    }
}
