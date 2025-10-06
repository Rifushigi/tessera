package org.tessera.domain.infrastructure.reader;

import org.apache.poi.xwpf.usermodel.*;
import org.tessera.domain.model.CandidateRecord;
import org.tessera.domain.model.Placeholder;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

public class WordDocumentWriter {

    private final Path templatePath;
    private final Set<Placeholder> placeholders;

    public WordDocumentWriter(Path templatePath, Set<Placeholder> placeholders) {
        this.templatePath = templatePath;
        this.placeholders = placeholders;
    }

    /**
     * Generates a single personalised document for a given candidate record.
     *
     * @param record       The candidate record containing the data for replacement.
     * @param outputStream The output stream to which the new document will be written.
     * @throws IOException if an error occurs during the file I/O.
     *
     */
    public void generateDocument(CandidateRecord record, OutputStream outputStream) throws IOException {

        try (InputStream templateStream = Files.newInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(templateStream)) {

            Map<String, String> replacements = getReplacementsMap(record);
            replacePlaceholders(document, replacements);

            document.write(outputStream);
        }
    }

    /**
     * Replaces all the placeholders in the document (paragraphs and tables)
     * with the values from the replacements map.
     * This implementation is optimised for placeholders that appear only once... For now :)
     *
     */
    private void replacePlaceholders(XWPFDocument document, Map<String, String> replacements) {
        // Replace in main paragraphs
        for (XWPFParagraph p : document.getParagraphs()) {
            replaceInParagraph(p, replacements);
        }

        //Replace in tables
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        replaceInParagraph(p, replacements);
                    }
                }
            }
        }
    }

    private void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> replacements) {

        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (text == null || text.isBlank()) {
                continue;
            }

            for (Placeholder placeholder : placeholders) {
                if (text.contains(placeholder.fullText())) {
                    String replacementValue = replacements.getOrDefault(placeholder.varName(), "");

                    // replace the text in the run
                    text = text.replace(placeholder.fullText(), replacementValue);
                    run.setText(text, 0);

                    // apply specific formatting based on the var name
                    applySpecificFormatting(run, placeholder.varName());
                }
            }
        }
    }

    private void applySpecificFormatting(XWPFRun run, String varName) {

        if (varName.equalsIgnoreCase("fullName")) {
            run.setFontFamily("Lucida Calligraphy");
            run.setFontSize(22);
        } else if (varName.equalsIgnoreCase("date")) {
            run.setFontSize(16);
        }
    }

    private Map<String, String> getReplacementsMap(CandidateRecord record) {

        Map<String, String> replacements = new HashMap<>();
        for (Placeholder p : placeholders) {
            String value = record.getValue(p.varName());
            replacements.put(p.varName(), Objects.requireNonNullElse(value, ""));
        }

        return replacements;
    }
}
