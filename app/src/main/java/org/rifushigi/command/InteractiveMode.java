package org.rifushigi.command;

import org.rifushigi.util.AnsiColors;
import org.rifushigi.util.FileValidator;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 * A helper class to guide the user through the input process in interactive mode.
 */
public class InteractiveMode {

    private final Scanner scanner;

    public InteractiveMode(Scanner scanner) {
        this.scanner = scanner;
    }

    /**
     * Prompts the user for file paths and returns them in a Result object.
     * @return An object containing the validated file paths.
     */
    public Result collectInput() {
        List<Path> templatePaths = promptForFilePaths("template", ".docx");
        Path dataPath = promptForFilePath("data", ".xlsx");
        Path outputDirectory = promptForDirectoryPath("output", ".");

        return new Result(templatePaths, dataPath, outputDirectory);
    }

    private Path promptForFilePath(String fileType, String extension) {
        while (true) {
            System.out.printf("Enter the path to the %s file (e.g., %s): %n", fileType, "template" + extension);
            System.out.print(AnsiColors.colored(AnsiColors.GREEN, ">> "));
            String input = scanner.nextLine().trim();
            Path path = Paths.get(input);

            if (FileValidator.fileExistsAndIsReadable(path) && FileValidator.hasExtension(path, extension)) {
                return path;
            } else {
                System.err.println(AnsiColors.colored(AnsiColors.RED, "Invalid path or file extension. Please try again."));
            }
        }
    }

    private List<Path> promptForFilePaths(String fileType, String extension) {
        List<Path> paths = new ArrayList<>();
        System.out.printf("Enter the path to the %s file (e.g., %s) or press Enter to finish: %n", fileType, "template" + extension);
        while (true) {
            System.out.print(AnsiColors.colored(AnsiColors.GREEN, ">> "));
            String input = scanner.nextLine().trim();

            if (input.isEmpty()) {
                if (paths.isEmpty()) {
                    System.err.println(AnsiColors.colored(AnsiColors.RED, "You must provide at least one template file."));
                } else {
                    break; // Exit the loop when done
                }
            } else {
                Path path = Paths.get(input);
                if (FileValidator.fileExistsAndIsReadable(path) && FileValidator.hasExtension(path, extension)) {
                    paths.add(path);
                    System.out.printf("  Added template: %s%n", path);
                    System.out.print("Enter another path or press Enter to finish: ");
                } else {
                    System.err.println(AnsiColors.colored(AnsiColors.RED, "Invalid path or file extension. Please try again."));
                }
            }
        }
        return paths;
    }

    private Path promptForDirectoryPath(String dirType, String defaultPath) {
        while (true) {
            System.out.printf("Enter the path for the %s directory [%s]: %n", dirType, defaultPath);
            System.out.print(AnsiColors.colored(AnsiColors.GREEN, ">> "));
            String input = scanner.nextLine().trim();
            Path path = input.isEmpty() ? Paths.get(defaultPath) : Paths.get(input);

            if (FileValidator.createDirectoryIfNotExists(path)) {
                return path;
            } else {
                System.err.println(AnsiColors.colored(AnsiColors.RED, "Invalid path or could not create directory. Please try again."));
            }
        }
    }

    /**
     * A record to hold the results of the interactive input collection.
     * Now correctly uses a List<Path> for templatePaths.
     */
    public record Result(List<Path> templatePaths, Path dataPath, Path outputDirectory) {}
}
