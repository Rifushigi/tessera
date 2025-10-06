package org.rifushigi.domain.model;

/**
 * Represents a variable placeholder found within a document template.
 *
 * @param varName The name of the variable, e.g "fullName".
 * @param fullText The full text of the placeholder including delimiters, e.g ${fullName}.
 * */
public record Placeholder(String varName, String fullText) {

    public Placeholder{
        if (varName == null || varName.isBlank()){
            throw new IllegalArgumentException("Variable name cannot be blank or null");
        }

        if (fullText == null || fullText.isBlank()){
            throw new IllegalArgumentException("Full text cannot be null or blank.");
        }
    }
}
