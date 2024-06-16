package example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class sample {
    public static void main(String[] args) {
        String excelFilePath = System.getProperty("user.dir") + "\\files\\Members.xlsx";
        File excelFile = new File(excelFilePath);
        
        try (FileInputStream fis = new FileInputStream(excelFile); 
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            
            XSSFSheet sheet = workbook.getSheet("one");
            int rows = sheet.getLastRowNum();
            int cols = sheet.getRow(1).getLastCellNum();

            for (int r = 0; r <= rows; r++) {
                XSSFRow row = sheet.getRow(r);
                for (int c = 0; c < cols; c++) {
                    XSSFCell cell = row.getCell(c);
                    CellType cellType = cell.getCellType();

                    switch (cellType) {
                        case STRING:
                            System.out.println(cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        default:
                            System.out.println("Unknown cell type");
                            break;
                    }
                }
            }
        } catch (FileNotFoundException e) {
            System.err.println("Excel file not found: " + e.getMessage());
        } catch (IOException e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
        }
    }
}
