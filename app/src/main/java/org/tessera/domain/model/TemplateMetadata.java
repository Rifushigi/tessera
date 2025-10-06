package org.tessera.domain.model;

import java.nio.file.Path;
import java.util.Collections;
import java.util.Set;

/**
 * Store metadata and detected placeholders from a document template
 *
 * @param templatePath The file system path to the .docx template file.
 * @param placeholders A set of unique {@link Placeholder} object found in the template.
 * */
public record TemplateMetadata(Path templatePath, Set<Placeholder> placeholders) {

    public TemplateMetadata{
        if (templatePath == null){
            throw new IllegalArgumentException("Template path cannot be null");
        }

        if (placeholders == null){
            throw new IllegalArgumentException("Placeholders set cannot be null");
        }
    }

    /**
     * @return An immutable view of the placeholders to prevent external modification
     * */
    public Set<Placeholder> getPlaceholders(){
        return Collections.unmodifiableSet(placeholders);
    }
}
