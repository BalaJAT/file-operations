package learning.java;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExelReader {

    public static void main(String[] args) {
        try {
            // Create a file input stream for the Excel file
            FileInputStream fileInputStream = new FileInputStream("output.xlsx");

            // Create a workbook from the input stream
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Get the first sheet in the workbook
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each row in the sheet
            for (Row row : sheet) {
                // Iterate through each cell in the row
                for (Cell cell : row) {
                    // Print the cell value to the console
                    System.out.print(cell + "\t");
                }
                System.out.println(); // Move to the next line for a new row
            }

            // Close the workbook and input stream to release resources
            workbook.close();
            fileInputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
}
