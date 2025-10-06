package org.rifushigi.command;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.rifushigi.domain.service.GenerationService;
import org.rifushigi.util.AnsiColors;
import org.rifushigi.util.FileValidator;
import picocli.CommandLine;

import java.io.File;
import java.nio.file.Path;
import java.util.Scanner;
import java.util.concurrent.Callable;

public class GenerateCommand implements Callable<Integer> {
    private static final Logger logger = LoggerFactory.getLogger(GenerateCommand.class);

    @CommandLine.Option(names = {"-t", "--template"}, description = "Path to the Word template (.docx).")
    private File templateFile;

    @CommandLine.Option(names = {"-d", "--data"}, description = "Path to the Excel data file (.xlsx).")
    private File dataFile;

    @CommandLine.Option(names = {"-o", "--output"}, defaultValue = "./output", description = "Output directory for generated documents.")
    private File outputDirectory;

    @CommandLine.Option(names = {"-i", "--interactive"}, description = "Run in interactive mode, prompting for input.")
    private boolean interactiveMode;

    @Override
    public Integer call() {
        if (interactiveMode) {
            System.out.println(AnsiColors.colored(AnsiColors.CYAN, "🚀 Starting Tessera in interactive mode..."));
            try (Scanner scanner = new Scanner(System.in)) {
                InteractiveMode interactive = new InteractiveMode(scanner);
                InteractiveMode.Result result = interactive.collectInput();
                return runGeneration(result.templatePath(), result.dataPath(), result.outputDirectory());
            }
        } else {
            // Validate command-line arguments
            if (templateFile == null || dataFile == null) {
                logger.error(AnsiColors.colored(AnsiColors.RED, "Missing required command-line arguments. Use --help for details or run with --interactive."));
                return 1;
            }
            return runGeneration(templateFile.toPath(), dataFile.toPath(), outputDirectory.toPath());
        }
    }

    private Integer runGeneration(Path templatePath, Path dataPath, Path outputDirectory) {
        // Validation moved from GenerationService to the command level
        if (!FileValidator.fileExistsAndIsReadable(templatePath) || !FileValidator.hasExtension(templatePath, ".docx")) {
            logger.error(AnsiColors.colored(AnsiColors.RED, "Invalid template file: " + templatePath));
            return 1;
        }
        if (!FileValidator.fileExistsAndIsReadable(dataPath) || !FileValidator.hasExtension(dataPath, ".xlsx")) {
            System.err.println(AnsiColors.colored(AnsiColors.RED, "Invalid data file: " + dataPath));
            return 1;
        }
        if (!FileValidator.createDirectoryIfNotExists(outputDirectory)) {
            System.err.println(AnsiColors.colored(AnsiColors.RED, "Could not create output directory: " + outputDirectory));
            return 1;
        }

        try {
            GenerationService service = new GenerationService(templatePath, dataPath, outputDirectory);
            service.generate();
            return 0;
        } catch (Exception e) {
            logger.error(AnsiColors.colored(AnsiColors.RED, "An unexpected error occurred during generation: " + e.getMessage()));
            return 1;
        }
    }
}
