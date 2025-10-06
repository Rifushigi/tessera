package org.tessera.domain.service;

import org.tessera.domain.infrastructure.reader.ExcelReader;
import org.tessera.domain.infrastructure.reader.WordDocumentWriter;
import org.tessera.domain.model.CandidateRecord;
import org.tessera.domain.model.Placeholder;
import org.tessera.util.AnsiColors;
import org.tessera.util.FileValidator;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class GenerationService {

    private final Path templatePath;
    private final Path dataPath;
    private final Path outputDir;

    public GenerationService(Path templatePath, Path dataPath, Path outputDir) {
        this.templatePath = templatePath;
        this.dataPath = dataPath;
        this.outputDir = outputDir;
    }


    /**
     * Executes the main document generation workflow.
     *
     * @throws IOException if there's an error with file I/O.
     */
    public void generate() throws IOException {
        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "📘 Scanning template for placeholders..."));
        PlaceholderService placeholderService = new PlaceholderService();
        Set<Placeholder> placeholders = placeholderService.findPlaceholders(templatePath);
        if (placeholders.isEmpty()) {
            System.err.println(AnsiColors.colored(AnsiColors.RED, "❌ Error: No placeholders found in the template. Please check your template file."));
            return;
        }
        System.out.printf("→ Found: %s%n", placeholders.stream().map(Placeholder::fullText).toList());

        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "📗 Reading Excel file..."));
        ExcelReader excelReader = new ExcelReader();
        Map<String, List<CandidateRecord>> allRecords = excelReader.readData(dataPath);
        if (allRecords.isEmpty()) {
            System.err.println(AnsiColors.colored(AnsiColors.RED, "❌ Error: No data found in the Excel file."));
            return;
        }

        System.out.printf("→ Detected sheets: %s%n", allRecords.keySet());
        int totalRecords = allRecords.values().stream().mapToInt(List::size).sum();
        System.out.printf("→ %d records total%n", totalRecords);

        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "⚙️ Mapping columns..."));
        for (Placeholder p : placeholders) {
            String varName = p.varName();
            if (allRecords.values().stream().anyMatch(list -> !list.isEmpty() && list.getFirst().getValue(varName) != null)) {
                System.out.printf("  %s → %s%n", p.fullText(), varName);
            } else {
                System.err.printf(AnsiColors.colored(AnsiColors.YELLOW, "Warning: Placeholder %s not found in Excel data.%n"), p.fullText());
            }
        }

        System.out.println(AnsiColors.colored(AnsiColors.CYAN, "Generating documents..."));
        WordDocumentWriter documentWriter = new WordDocumentWriter(templatePath, placeholders);
        int filesCreated = 0;

        for (Map.Entry<String, List<CandidateRecord>> entry : allRecords.entrySet()) {
            String sheetName = entry.getKey();
            List<CandidateRecord> records = entry.getValue();

            Path sheetOutputDir = outputDir.resolve(sheetName);
            if (!FileValidator.createDirectoryIfNotExists(sheetOutputDir)) {
                System.err.println(AnsiColors.colored(AnsiColors.RED, "Failed to create output directory: " + sheetOutputDir));
                continue; // Continue to the next sheet
            }

            for (CandidateRecord record : records) {
                String fileName = record.getValue("fullName") + ".docx";
                Path outputPath = sheetOutputDir.resolve(fileName);

                try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
                    documentWriter.generateDocument(record, fos);
                    filesCreated++;
                } catch (IOException e) {
                    System.err.printf(AnsiColors.colored(AnsiColors.RED, "Error generating document for %s: %s%n"), record.getValue("fullName"), e.getMessage());
                }
            }
        }
        System.out.printf(AnsiColors.colored(AnsiColors.GREEN, "%d files created in %s (organized by sheet)%n"), filesCreated, outputDir.toAbsolutePath());
    }
}
