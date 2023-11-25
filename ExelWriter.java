package learning.java;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExelWriter {
	public static void main(String[] args) {
        try {
            // Create a new workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            // Define column headers
            String[] headers = {"Name", "Age", "Email"};

            // Create the header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Data to be written
            Object[][] data = {
                    {"John Doe", 30, "john@test.com"},
                    {"Jane Doe", 28, "jane@test.com"},
                    {"Bob Smith", 35, "jacky@example.com"},
                    {"Swapnil", 37, "joe@example.com"}
            };

            // Write data to the sheet
            for (int rowIndex = 0; rowIndex < data.length; rowIndex++) {
                Row row = sheet.createRow(rowIndex + 1); // Start from the second row (index 1)
                for (int columnIndex = 0; columnIndex < data[rowIndex].length; columnIndex++) {
                    Cell cell = row.createCell(columnIndex);
                    if (data[rowIndex][columnIndex] instanceof String) {
                        cell.setCellValue((String) data[rowIndex][columnIndex]);
                    } else if (data[rowIndex][columnIndex] instanceof Integer) {
                        cell.setCellValue((Integer) data[rowIndex][columnIndex]);
                    }
                    // Add more conditions for other data types if necessary
                }
            }

            // Write the workbook content to a file
            try (FileOutputStream fileOut = new FileOutputStream("output.xlsx")) {
                workbook.write(fileOut);
            }

            // Close the workbook to release resources
            workbook.close();

            System.out.println("Excel file has been created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
