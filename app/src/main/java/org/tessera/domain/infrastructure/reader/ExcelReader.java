package org.tessera.domain.infrastructure.reader;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.tessera.domain.model.CandidateRecord;
import org.tessera.domain.model.Placeholder;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

public class ExcelReader {

    /**
     * Reads data from all sheets of an Excel file and returns a map where keys
     * are sheet names and values are the lists of {@link Placeholder} objects.
     *
     * @param excelFilePath The path to the Excel file.
     * @return A map of sheet names to lists to candidate records.
     * @throws IOException if there's an error reading the file.
     * */
    public Map<String, List<CandidateRecord>> readData(Path excelFilePath) throws IOException {
        if (!Files.exists(excelFilePath) || !Files.isReadable(excelFilePath)) {
            throw new IOException("Excel file not found or is not readable: " + excelFilePath);
        }

        Map<String, List<CandidateRecord>> dataBySheet = new LinkedHashMap<>();

        try(InputStream is = Files.newInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(is)){
            for (int i = 0; i < workbook.getNumberOfSheets(); i++){
                Sheet sheet = workbook.getSheetAt(i);
                if (sheet.getPhysicalNumberOfRows() > 1){
                    List<CandidateRecord> records = parseSheet(sheet);
                    dataBySheet.put(sheet.getSheetName(), records);
                }
            }
        }

        return dataBySheet;
    }

    /**
     * Parses a single sheet from the workbook.
     *
     * @param sheet The sheet to parse.
     * @return A list of candidate records for the sheet.
     * */
    private List<CandidateRecord> parseSheet(Sheet sheet){
        List<CandidateRecord> records = new ArrayList<>();
        DataFormatter dataFormatter = new DataFormatter();

        // I'm assuming the first row is the header
        // As it should be :)
        Row headerRow = sheet.getRow(0);
        if(headerRow == null){
            return records;
        }

        List<String> headers = new ArrayList<>();
        for (Cell cell: headerRow){
            headers.add(dataFormatter.formatCellValue(cell));
        }

        // Iterate over the remaining rows
        for (int i = 1; i <= sheet.getLastRowNum(); i++){
            Row dataRow = sheet.getRow(i);
            if (dataRow == null){
                continue;
            }

            Map<String, String> rowData = new HashMap<>();
            boolean hasData = false;

            for (int j = 0; j < headers.size(); j++){
                Cell cell = dataRow.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                String cellValue = (cell == null) ? "" : dataFormatter.formatCellValue(cell);
                rowData.put(headers.get(j), cellValue);

                if(!cellValue.isBlank()){
                    hasData = true;
                }
            }
            if (hasData){
                records.add(new CandidateRecord(rowData));
            }
        }
        return records;
    }
}
