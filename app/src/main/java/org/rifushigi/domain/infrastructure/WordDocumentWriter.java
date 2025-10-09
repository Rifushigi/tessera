package org.rifushigi.domain.infrastructure;

import org.apache.poi.xwpf.usermodel.*;
import org.rifushigi.domain.model.CandidateRecord;
import org.rifushigi.domain.model.Placeholder;

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
     * Generates a single personalised document (DOCX format) for a given candidate record.
     *
     * @param record       The candidate record containing the data for replacement.
     * @param outputStream The output stream to which the new document will be written.
     * @throws IOException if an error occurs during the file I/O.
     */
    public void generateDocument(CandidateRecord record, OutputStream outputStream) throws IOException {
        try (XWPFDocument document = createAndReplaceDocument(record)) {
            document.write(outputStream);
        }
    }

    /**
     * Creates a new document from the template and applies all placeholder replacements.
     */
    private XWPFDocument createAndReplaceDocument(CandidateRecord record) throws IOException {
        XWPFDocument document;
        try (InputStream templateStream = Files.newInputStream(templatePath)) {
            document = new XWPFDocument(templateStream);
        }

        Map<String, String> replacements = getReplacementsMap(record);
        replacePlaceholders(document, replacements);

        return document;
    }

    /**
     * Replaces all the placeholders in the document
     * with the values from the replacements map.
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

    /**
     * Rewritten to handle fragmented placeholders across multiple XWPFRun elements.
     * It performs replacement on the paragraph's concatenated text, then deletes
     * the old runs and creates a new run with the replaced text.
     */
    private void replaceInParagraph(XWPFParagraph paragraph, Map<String, String> replacements) {
        String paragraphText = paragraph.getText();

        for (Placeholder placeholder : placeholders) {
            String fullTag = placeholder.fullText();

            if (paragraphText.contains(fullTag)) {
                String replacementValue = replacements.getOrDefault(placeholder.varName(), "");

                String newParagraphText = paragraphText.replace(fullTag, replacementValue);

                // delete all existing runs in the paragraph (backward)
                int numRuns = paragraph.getRuns().size();
                for (int i = numRuns - 1; i >= 0; i--) {
                    paragraph.removeRun(i);
                }

                // create a new run with the replaced text
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(newParagraphText, 0);

                applySpecificFormatting(newRun, placeholder.varName());

                return; // Exit after replacing one placeholder to simplify the run management
            }
        }
    }

    private void applySpecificFormatting(XWPFRun run, String varName) {

        if (varName.equalsIgnoreCase("FULL NAME")) {
            run.setFontFamily("Lucida Calligraphy");
            run.setFontSize(22);
        } else if (varName.equalsIgnoreCase("DATE")) {
            run.setFontFamily("Swis721 Th TL");
            run.setFontSize(16);
        }
    }

    private Map<String, String> getReplacementsMap(CandidateRecord record) {

        Map<String, String> replacements = new HashMap<>();
        for (Placeholder p : placeholders) {
            String varName = p.varName();
            String value = record.getValue(varName);
            replacements.put(varName, Objects.requireNonNullElse(value, ""));

            // Override the DATE variable with the user-requested value
            if (varName.equalsIgnoreCase("DATE")) {
                replacements.put(varName, "20th September 2025");
            }
        }

        return replacements;
    }
}
