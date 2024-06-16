package example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
    public static void main(String[] args) {
        String excelFilePath = System.getProperty("user.dir") + "\\Resources\\Members.xlsx";
        File excelFile = new File(excelFilePath);
        
        if (!excelFile.exists()) {
            System.err.println("Excel file not found at path: " + excelFilePath);
            return;
        }

        try (FileInputStream fis = new FileInputStream(excelFile); 
             XSSFWorkbook workbook = new XSSFWorkbook(fis);
             Connection conn = DataBaseUtils.getConnection()) {
            
            XSSFSheet sheet = workbook.getSheet("one");
            if (sheet == null) {
                System.err.println("Sheet 'one' not found in the workbook");
                return;
            }
            
            int rows = sheet.getLastRowNum();
            int cols = sheet.getRow(0).getLastCellNum();  // Assuming first row exists and is used for column count
            
            String sql = "INSERT INTO members (name, id, email) VALUES (?, ?, ?)";
            try (PreparedStatement stmt = conn.prepareStatement(sql)) {
                for (int r = 1; r <= rows; r++) {  // Assuming first row is header
                    XSSFRow row = sheet.getRow(r);
                    if (row == null) {
                        continue; // Skip null rows
                    }
                    String name = row.getCell(0).getStringCellValue();
                    double id = row.getCell(1).getNumericCellValue();
                    String email = row.getCell(2).getStringCellValue();

                    //get add
                    stmt.setString(1, name);
                    stmt.setInt(2, (int) id);
                    stmt.setString(3, email);
                    stmt.addBatch();
                }
                stmt.executeBatch();  // Execute batch insert
            }

            System.out.println("Data inserted successfully.");
        } catch (FileNotFoundException e) {
            System.err.println("Excel file not found: " + e.getMessage());
        } catch (IOException e) {
            System.err.println("Error reading Excel file: " + e.getMessage());
        } catch (SQLException e) {
            System.err.println("Database error: " + e.getMessage());
        }
    }
}
