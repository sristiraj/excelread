package excelread;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;



public class App{

    public void readXlsFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new HSSFWorkbook(fis)) {
    
            // Get the first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);
            
            // Iterate over each row in the sheet
            for (Row row : sheet) {
                // Iterate over each cell in the row
                for (Cell cell : row) {
                    // Get the cell value and print it
                    CellType cellType = cell.getCellType();
                    if (cellType == CellType.STRING) {
                        System.out.print(cell.getStringCellValue() + "\t");
                    } else if (cellType == CellType.NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + "\t");
                    } else if (cellType == CellType.BOOLEAN) {
                        System.out.print(cell.getBooleanCellValue() + "\t");
                    } else if (cellType == CellType.BLANK) {
                        System.out.print("\t");
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    
    public void readXlsxFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
    
            // Get the first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);
    
            // Iterate over each row in the sheet
            for (Row row : sheet) {
                // Iterate over each cell in the row
                for (Cell cell : row) {
                    // Get the cell value and print it
                    CellType cellType = cell.getCellType();
                    if (cellType == CellType.STRING) {
                        System.out.print(cell.getStringCellValue() + "\t");
                    } else if (cellType == CellType.NUMERIC) {
                        System.out.print(cell.getNumericCellValue() + "\t");
                    } else if (cellType == CellType.BOOLEAN) {
                        System.out.print(cell.getBooleanCellValue() + "\t");
                    } else if (cellType == CellType.BLANK) {
                        System.out.print("\t");
                    }
                }
                System.out.println();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    public List<List<String>> getEmailsFromExcel(String filePath, int[] columnIndexes, int[] mainColumnIndexes) {
        
        List<List<String>> rowList = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new HSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                List<String> mailList = new ArrayList<>();
                List<String> emailList = new ArrayList<>();
                
                for (int columnIndex : columnIndexes) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null) {
                        CellType cellType = cell.getCellType();
                        if (cellType == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            List<String> email = identifyEmails(cellValue);
                            emailList.addAll(email);
                            
                        }
                    }
                }

                for (int columnIndex : mainColumnIndexes) {
                    Cell cell = row.getCell(columnIndex);
                    if (cell != null) {
                        CellType cellType = cell.getCellType();
                        if (cellType == CellType.STRING) {
                            String cellValue = cell.getStringCellValue();
                            mailList.add(cellValue);
                        }
                        if (cellType == CellType.NUMERIC) {
                            Double cellValue = cell.getNumericCellValue();
                            int cellIntValue = (int) Math.floor(cellValue);
                            mailList.add(Integer.valueOf(cellIntValue).toString());
                            
                        }
                    }
                }
                mailList.add(String.join(";", emailList));
                
                rowList.add(mailList);
                
                
            }
            

        } catch (IOException e) {
            e.printStackTrace();
        }

        return rowList;
    }

    public List<String> identifyEmails(String input) {
        List<String> emails = new ArrayList<>();

        // Regular expression pattern for email matching
        String regex = "\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}\\b";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(input);

        // Find and identify email addresses
        while (matcher.find()) {
            String email = matcher.group();
            emails.add(email);
        }

        return emails;
    }

    private boolean isValidEmail(String email) {
        // Simple email validation regex
        String emailRegex = "^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+$";
        return email.matches(emailRegex);
    }

    public void saveListToExcel(List<List<String>> data, String filePath) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");
    
        // Create header row
        // Row headerRow = sheet.createRow(0);
        // CellStyle headerCellStyle = workbook.createCellStyle();
        // Font headerFont = workbook.createFont();
        // headerFont.setBold(true);
        // headerCellStyle.setFont(headerFont);
    
        // Write data to the sheet
        int rowIndex = 0;
        for (List<String> row : data) {
            Row sheetRow = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            for (String cellValue : row) {
                Cell cell = sheetRow.createCell(cellIndex++);
                cell.setCellValue(cellValue);
            }
        }

        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    
    public static void main(String[] args){
        System.out.println("Started parsing...");
        if (args.length > 0) {
            // Access individual command-line arguments
            for (int i = 0; i < args.length; i++) {
                System.out.println("Argument " + (i + 1) + ": " + args[i]);
            }
        } else {
            System.out.println("No command-line arguments provided.");
        }
        String filePath = args[0];

        int[] columnIndexes = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}; // Column indexes (0-based) to extract email strings
        int[] mainColumnIndexes = {0,1,2};
        App excelReader = new App();
        List<List<String>> emails = excelReader.getEmailsFromExcel(filePath, columnIndexes, mainColumnIndexes);

        excelReader.saveListToExcel(emails, "output.xls");
    }
}

