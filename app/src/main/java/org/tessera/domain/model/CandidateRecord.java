package org.tessera.domain.model;

import java.util.Map;
import java.util.Objects;

/**
* Represents a single record of a candidate's data read from an Excel sheet.
* The data is stored as a map of column headers (keys) to cell values (values).
*
* @param data A map that represents the candidate's data.
*/
public record CandidateRecord(Map<String, String> data) {

    public CandidateRecord{
        Objects.requireNonNull(data, "Data map can not be null");
    }

    /**
     * Gets the value for a given header (case-insensitive)
     *
     * @param header The column header to retrieve the value for.
     * @return The value or null if the header is not found.
     */
    public String getValue(String header){
        if(header == null){
            return null;
        }

        return data.entrySet().stream()
                .filter(entry -> entry.getKey().equalsIgnoreCase(header))
                .findFirst()
                .map(Map.Entry::getValue)
                .orElse(null);
    }
}
