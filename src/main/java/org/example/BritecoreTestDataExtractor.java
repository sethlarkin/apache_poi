package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.*;
import java.util.ArrayList;
import java.util.List;

public class BritecoreTestDataExtractor {

    private final String filePath;
    //ADD LOGGER FOR BETTER ERROR MONITORING

    public BritecoreTestDataExtractor(String filePath) {
        this.filePath = filePath;
    }

    public List<String> getSheetNames() {
        List<String> sheetNames = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(this.filePath); Workbook workbook = new XSSFWorkbook(fis);) {
            int numSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < numSheets; i++) {
                sheetNames.add(workbook.getSheetName(i));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return sheetNames;
    }


    public List<String> getColumnNames(String sheetName) {
        List<String> columnNames = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(this.filePath); Workbook workbook = new XSSFWorkbook(fis);) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet != null) {
                Row headerRow = sheet.getRow(0); // names must be in first row
                if (headerRow != null) {
                    for (Cell cell : headerRow) {
                        columnNames.add(cell.getStringCellValue());
                    }
                } else {
                    System.err.println("Header row is empty");
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return columnNames;
    }

    public List<String> getRowData(String sheetName) {
        List<String> rowData = new ArrayList<>();
        return rowData;
    }

    public List<Map<String, String>> getRowDataById(int quoteId, String sheetName) {
        List<Map<String, String>> matchingRows = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(this.filePath); Workbook workbook = new XSSFWorkbook(fis);) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet != null) {
                int quoteIdColumnIndex = -1;
                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    for (Cell cell : headerRow) {
                        if (cell.getStringCellValue().equalsIgnoreCase("quote id")) {
                            quoteIdColumnIndex = cell.getColumnIndex();
                            break;
                        }
                    }
                    if (quoteIdColumnIndex != -1) {
                        for (Row row : sheet) {
                            if (row.getRowNum() == 0) { //skipping header row
                                continue;
                            }
                            Cell cell = row.getCell(quoteIdColumnIndex);
                            if (cell != null && cell.getCellType() == CellType.NUMERIC && cell.getNumericCellValue() == quoteId) {
                                Map<String, String> rowData = new LinkedHashMap<>();
                                for (Cell c : row) {
                                    String columnName = headerRow.getCell(c.getColumnIndex()).getStringCellValue();
                                    rowData.put(columnName, c.toString());
                                }
                                matchingRows.add(rowData);
                            }
                        }
                    } else {
                        System.err.println("No Quote Id column found!");
                    }
                } else {
                    System.err.println("No Header row found!");

                }
            } else {
                System.err.println("Sheet not found!");

            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return matchingRows;
    }

    public Set<String> getAvailablePolicyTypes(int quoteId, String sheetName) {
        System.out.println("Getting available policy types: ");
        Set<String> policyTypes = new HashSet<>();
        List<Map<String, String>> matchingRows = getRowDataById(quoteId, sheetName);
        for (Map<String, String> row : matchingRows) {
            String policyType = row.get("Available Policy Type");
            if (policyType != null && !policyType.isEmpty()) {
                policyTypes.add(policyType);
            }
        }
        return policyTypes;
    }

    public Map<String, Object> createPolicyTestData(int quoteId, List<String> sheets) {
        Map<String, Object> policyTestData = new HashMap<>();
        policyTestData.put("quoteId", quoteId);
        List<String> availablePolicyTypes = new ArrayList<>();

        for (String sheetName : sheets) {
            try (FileInputStream fis = new FileInputStream(this.filePath); Workbook workbook = new XSSFWorkbook(fis);) {
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    System.err.println("Sheet " + sheetName + " not found");
                    continue;
                }
                Row headerRow = sheet.getRow(0);
                if (headerRow == null) {
                    System.err.println("Header row missing in sheet " + sheetName);
                }

                boolean quoteIdFound = false;
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    Cell quoteIdCell = row.getCell(0);
                    if (quoteIdCell == null || quoteIdCell.getCellType() != CellType.NUMERIC) {
                        continue;
                    }

                    if ((int) quoteIdCell.getNumericCellValue() == quoteId) {
                        quoteIdFound = true;
                        for (Cell cell : row) {
                            String columnName = headerRow.getCell(cell.getColumnIndex()).getStringCellValue();
                            if (columnName == null || columnName.isEmpty()) {
                                continue;
                            }
                            try {
                                if (columnName.equalsIgnoreCase("available policy type")) {
                                    availablePolicyTypes.add((cell.getStringCellValue()));
                                } else {
//                                    String value = cell.getCellType() == CellType.STRING ? cell.getStringCellValue() :
//                                            cell.getCellType() == CellType.NUMERIC ? (cell.getNumericCellValue() % 1 == 0 ? String.valueOf((int) cell.getNumericCellValue()) : String.valueOf(cell.getNumericCellValue())) : "";
                                    String value = cell.toString();
                                    policyTestData.put(columnName, value);
                                }
                            } catch (Exception e) {
                                System.err.println("Error reading cell in sheet: " + sheetName + " at row: " + (row.getRowNum() + 1) + " and column: " + (cell.getColumnIndex() + 1) + " error: " + e.getMessage());
                            }
                        }
                    }
                    if (!quoteIdFound) {
                        System.err.println("Quote ID not found in " + sheetName);
                    }
                }
                policyTestData.put("Available Policy Types", availablePolicyTypes);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return policyTestData;
    }

}
