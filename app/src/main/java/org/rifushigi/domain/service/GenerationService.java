package org.rifushigi.domain.service;

import org.rifushigi.domain.infrastructure.ExcelReader;
import org.rifushigi.domain.infrastructure.WordDocumentWriter;
import org.rifushigi.domain.model.CandidateRecord;
import org.rifushigi.domain.model.Placeholder;
import org.rifushigi.util.AnsiColors;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class GenerationService {

    private final List<Path> templatePaths;
    private final Path dataPath;
    private final Path outputDir;

    private static final String DOCX_SUBDIR = "certificates_09_2025";

    public GenerationService(List<Path> templatePaths, Path dataPath, Path outputDir) {
        this.templatePaths = templatePaths;
        this.dataPath = dataPath;
        this.outputDir = outputDir;
    }

    /**
     * Executes the main document generation workflow.
     *
     * @throws IOException if there's an error with file I/O.
     */
    public void generate() throws IOException {
        // Read Excel Data (only once)
        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "Reading Excel file..."));
        ExcelReader excelReader = new ExcelReader();
        Map<String, List<CandidateRecord>> allRecords = excelReader.readData(dataPath);
        if (allRecords.isEmpty()) {
            System.err.println(AnsiColors.colored(AnsiColors.RED, "Error: No data found in the Excel file."));
            return;
        }

        System.out.printf("Detected sheets: %s%n", allRecords.keySet());
        int totalRecords = allRecords.values().stream().mapToInt(List::size).sum();
        System.out.printf("%d records total%n", totalRecords);

        // Map sheet names to specific template paths for a single scan
        Map<String, Path> templateMap = createSheetTemplateMap();

        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "Generating documents..."));

        for (Map.Entry<String, List<CandidateRecord>> entry : allRecords.entrySet()) {
            String sheetName = entry.getKey();
            List<CandidateRecord> records = entry.getValue();

            Path specificTemplatePath = templateMap.get(sheetName);
            if (specificTemplatePath == null) {
                System.err.printf(AnsiColors.colored(AnsiColors.YELLOW, "%n️No template specified for sheet '%s'. Skipping this sheet.%n"), sheetName);
                continue;
            }

            System.out.printf(AnsiColors.colored(AnsiColors.CYAN, "%nProcessing sheet '%s' with template '%s'...%n"), sheetName, specificTemplatePath.getFileName());

            // Scan template for placeholders (once per template)
            PlaceholderService placeholderService = new PlaceholderService();
            Set<Placeholder> placeholders = placeholderService.findPlaceholders(specificTemplatePath);
            if (placeholders.isEmpty()) {
                System.err.println(AnsiColors.colored(AnsiColors.RED, "Error: No placeholders found in this template. Skipping."));
                continue;
            }

            // Map columns (only for the current template)
            logMapping(placeholders, records);

            WordDocumentWriter documentWriter = new WordDocumentWriter(specificTemplatePath, placeholders);
            String templateBaseName = specificTemplatePath.getFileName().toString().replace(".docx", "");

            // Define base output paths for DOCX, organized by template name
            // Example structure: outputDir/docx/diploma/Level 2
            Path docxBaseDir = outputDir.resolve(DOCX_SUBDIR).resolve(templateBaseName).resolve(sheetName);

            int filesCreatedDocx = 0;

            for (CandidateRecord record : records) {
                String baseFileName = record.getValue("FULL NAME");

                // generate DOCX
                Path docxPath = docxBaseDir.resolve(baseFileName + ".docx");
                if (ensureDirectoryExists(docxPath)) {
                    try (FileOutputStream fos = new FileOutputStream(docxPath.toFile())) {
                        documentWriter.generateDocument(record, fos);
                        filesCreatedDocx++;
                    } catch (IOException e) {
                        System.err.printf(AnsiColors.colored(AnsiColors.RED, "Error generating DOCX for %s: %s%n"), baseFileName, e.getMessage());
                    }
                }
            }

            System.out.printf(AnsiColors.colored(AnsiColors.GREEN,
                            "%d DOCX files created for sheet '%s' in %s%n"),
                    filesCreatedDocx, sheetName, docxBaseDir.toAbsolutePath());
        }
        System.out.println(AnsiColors.colored(AnsiColors.GREEN, "Generation complete for all templates..."));
    }

    /**
     * Helper method to ensure the parent directory of a file path exists.
     */
    private boolean ensureDirectoryExists(Path filePath) {
        File parentDir = filePath.getParent().toFile();
        if (!parentDir.exists() && !parentDir.mkdirs()) {
            System.err.printf(AnsiColors.colored(AnsiColors.RED, "Failed to create output directory: %s. Skipping file creation.%n"), parentDir.getAbsolutePath());
            return false;
        }
        return true;
    }

    /**
     * Maps specific sheet names to template file paths.
     */
    private Map<String, Path> createSheetTemplateMap() {
        Map<String, Path> templateMap = new java.util.HashMap<>();
        for (Path path : templatePaths) {
            String fileName = path.getFileName().toString().toLowerCase();
            if (fileName.contains("diploma")) {
                templateMap.put("Level 2", path);
            } else if (fileName.contains("certificate")) {
                templateMap.put("Level 1", path);
                templateMap.put("Kainos OAU", path);
            }
        }
        return templateMap;
    }

    /**
     * Logs the mapping of placeholders to Excel columns for user feedback.
     */
    private void logMapping(Set<Placeholder> placeholders, List<CandidateRecord> records) {
        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "Mapping columns..."));
        if (records.isEmpty()) {
            System.err.println(AnsiColors.colored(AnsiColors.YELLOW, "Cannot map columns: No records found in sheet."));
            return;
        }
        CandidateRecord firstRecord = records.getFirst();
        for (Placeholder p : placeholders) {
            String varName = p.varName();
            if (firstRecord.getValue(varName) != null) {
                System.out.printf("  %s -> %s%n", p.fullText(), varName);
            } else {
                System.err.printf(AnsiColors.colored(AnsiColors.YELLOW, "Warning: Placeholder %s not found in Excel data.%n"), p.fullText());
            }
        }
    }
}
