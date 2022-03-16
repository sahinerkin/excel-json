package com.yktech.ExcelJSON.Controllers;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;

import java.time.LocalDateTime;
import java.util.*;
import java.util.List;

@RestController
@RequestMapping("/api/v1/excel")
public class ExcelController {


    /**
     * Processes and Excel file sheet by sheet and row by row
     * and returns the contents in JSON format
     * @param file - input Excel file (MultipartFile)
     * @return response for the API call (ResponseEntity<HashMap<String, Object>>)
     */
    @RequestMapping(value = "/process-excel-file",
                    method = RequestMethod.POST,
                    consumes = MediaType.MULTIPART_FORM_DATA_VALUE,
                    produces = MediaType.APPLICATION_JSON_VALUE)


    public ResponseEntity<HashMap<String, Object>> processExcelFile(@RequestParam("file") MultipartFile file,
                                           HttpServletRequest request) {

        // Return error message if file doesn't exist
        if (file == null || file.isEmpty()) {
            return getError(HttpStatus.BAD_REQUEST, "Dosya mevcut değil veya okunamadı.", request);
        }

        Workbook workbook;

        // Return error message if a problem is encountered when opening the file as an Excel file
        try {
            // Get the workbook
            workbook = new XSSFWorkbook(file.getInputStream());
        } catch (Exception e) {
            return getError(HttpStatus.UNPROCESSABLE_ENTITY, "Dosya geçerli bir Excel dosyası değil.", request);
        }

        // Create the list of sheets
        List<LinkedHashMap<String, Object>> sheetList = new ArrayList<>();

        // Create a sheet iterator and iterate on all the sheets in the workbook
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        while (sheetIterator.hasNext()) {
            // Get the name and contents of the current sheet
            Sheet sheet = sheetIterator.next();
            LinkedHashMap<String, Object> sheetMap = new LinkedHashMap<>();
            ArrayList<LinkedHashMap<String, String>> contentList = getContentsForSheet(sheet);

            if (contentList == null)
                continue;

            sheetMap.put("name", sheet.getSheetName());
            sheetMap.put("contents", getContentsForSheet(sheet));

            // Add to the sheet list
            sheetList.add(sheetMap);
        }

        return ResponseEntity.ok(
                new HashMap<>() {{
                    put("sheets", sheetList);
                }});

    }


    /**
     * Gets a worksheet and returns its rows in a list
     * @param sheet - input worksheet (Sheet)
     * @return the content as a list of rows (ArrayList<LinkedHashMap<String, String>>)
     */
    private ArrayList<LinkedHashMap<String, String>> getContentsForSheet(Sheet sheet) {

        if (sheet.getPhysicalNumberOfRows() < 2)
            return null;

        // Initialize the data formatter
        DataFormatter formatter = new DataFormatter();

        // List of headers as strings
        List<String> headerList = getHeaders(sheet, formatter);

        // List of rows as LinkedHashMaps (key-value pair maps)
        ArrayList<LinkedHashMap<String, String>> rowList = new ArrayList<>();

        // Iterate on all the rows
        for (int i = 1; i < sheet.getLastRowNum()+1; i++) {

            // Get the row and create LinkedHashMap for the row
            Row row = sheet.getRow(i);
            LinkedHashMap<String, String> map = new LinkedHashMap<>();

            // Iterate on the cells of the row
            for (int j = 0; j < headerList.size(); j++) {
                // Get the cell and put the value in the row map
                Cell cell = row.getCell(j);
                map.put(headerList.get(j), formatter.formatCellValue(cell));
            }

            // Add the row map to the list
            rowList.add(map);
        }

        return rowList;
    }


    /**
     * Gets a worksheet and returns its headers as a list of strings
     * @param sheet - input worksheet (Sheet)
     * @param formatter - data formatter, if already initialized (DataFormatter)
     * @return the headers in a list (ArrayList<String>)
     */
    private ArrayList<String> getHeaders(Sheet sheet, DataFormatter formatter) {

        // If formatter is not given, initialize
        if (formatter == null)
            formatter = new DataFormatter();

        // Create a list for the headers
        ArrayList<String> headerList = new ArrayList<>();

        // Get the header row (first row - index 0)
        Row headerRow = sheet.getRow(0);

        // Iterate on the header row and add the values to the list
        for (int j = 0; j < headerRow.getLastCellNum(); j++) {
            headerList.add(formatter.formatCellValue(headerRow.getCell(j)));
        }

        return headerList;
    }


    /**
     * Returns error response with the given status and message info
     * @param status - http status code to return (HttpStatus)
     * @param message - error message to add (message)
     * @param request - received request (HttpServletRequest)
     * @return the response entity (ResponseEntity<HashMap<String, Object>>)
     */
    private ResponseEntity<HashMap<String, Object>> getError(HttpStatus status, String message, HttpServletRequest request) {
        return ResponseEntity.status(status).body(new LinkedHashMap<>() {{
            put("timestamp", LocalDateTime.now());
            put("status", status.value());
            put("error", message);
            put("path", request.getServletPath());
        }});
    }
}

