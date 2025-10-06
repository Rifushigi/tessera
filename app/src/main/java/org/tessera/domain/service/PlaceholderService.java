package org.tessera.domain.service;

import org.apache.poi.xwpf.usermodel.*;
import org.tessera.domain.model.Placeholder;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PlaceholderService {

    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\$\\{(.+?)}");

    /**
     * Finds all unique placeholders within a given Word document template.
     *
     * @param templatePath The path to the .docx template file.
     * @return A set of unique {@link Placeholder} objects.
     * @throws IOException if there is an error reading the file.
     */
    public Set<Placeholder> findPlaceholders(Path templatePath) throws IOException {

        Set<Placeholder> placeholders = new HashSet<>();

        if (!Files.exists(templatePath) || !Files.isReadable(templatePath)){
            throw new IOException("Template file not found or is not readable");
        }

        try (InputStream is = Files.newInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(is);) {

            // scan paragraphs
            for (XWPFParagraph p : document.getParagraphs()) {
                scanParagraph(p, placeholders);
            }

            // Scan tables
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            scanParagraph(p, placeholders);
                        }
                    }
                }
            }
        }

        return placeholders;
    }

    private void scanParagraph(XWPFParagraph p, Set<Placeholder> placeholders){

        StringBuilder paragraphText = new StringBuilder();
        for (XWPFRun run : p.getRuns()){
            String text = run.getText(0);
            if (text != null){
                paragraphText.append(text);
            }
        }

        Matcher matcher = PLACEHOLDER_PATTERN.matcher(paragraphText.toString());
        while (matcher.find()) {
            String fullText = matcher.group();
            String varName = matcher.group(1);
            placeholders.add(new Placeholder(varName, fullText));
        }
    }
}
