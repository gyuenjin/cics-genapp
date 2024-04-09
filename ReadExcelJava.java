import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {

    public static void main(String[] args) {
        // File paths for the Excel spreadsheets
        String filePath1 = "path/to/spreadsheet1.xlsx";
        String filePath2 = "path/to/spreadsheet2.xlsx";

        try {
            // Open and read the first spreadsheet
            FileInputStream fis1 = new FileInputStream(new File(filePath1));
            Workbook workbook1 = WorkbookFactory.create(fis1);
            calculateMeanAndStdDev(workbook1);

            // Open and read the second spreadsheet
            FileInputStream fis2 = new FileInputStream(new File(filePath2));
            Workbook workbook2 = WorkbookFactory.create(fis2);
            calculateMeanAndStdDev(workbook2);

            // Close the input streams
            fis1.close();
            fis2.close();
        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

    private static void calculateMeanAndStdDev(Workbook workbook) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            double sumJ = 0, sumK = 0;
            int rowCount = Math.min(sheet.getLastRowNum(), 499); // Limit to 500 rows

            // Sum up the values in columns J and K
            for (int j = 1; j <= rowCount; j++) {
                Row row = sheet.getRow(j);
                if (row != null) {
                    Cell cellJ = row.getCell(9); // Column J (0-indexed)
                    Cell cellK = row.getCell(10); // Column K (0-indexed)
                    if (cellJ != null && cellK != null) {
                        sumJ += cellJ.getNumericCellValue();
                        sumK += cellK.getNumericCellValue();
                    }
                }
            }

            // Calculate the mean
            double meanJ = sumJ / rowCount;
            double meanK = sumK / rowCount;

            // Calculate the standard deviation
            double sumDeviationJ = 0, sumDeviationK = 0;
            for (int j = 1; j <= rowCount; j++) {
                Row row = sheet.getRow(j);
                if (row != null) {
                    Cell cellJ = row.getCell(9); // Column J (0-indexed)
                    Cell cellK = row.getCell(10); // Column K (0-indexed)
                    if (cellJ != null && cellK != null) {
                        sumDeviationJ += Math.pow(cellJ.getNumericCellValue() - meanJ, 2);
                        sumDeviationK += Math.pow(cellK.getNumericCellValue() - meanK, 2);
                    }
                }
            }
            double stdDevJ = Math.sqrt(sumDeviationJ / rowCount);
            double stdDevK = Math.sqrt(sumDeviationK / rowCount);

            // Print the results
            System.out.println("Sheet: " + sheet.getSheetName());
            System.out.println("Mean of Column J: " + meanJ);
            System.out.println("Standard Deviation of Column J: " + stdDevJ);
            System.out.println("Mean of Column K: " + meanK);
            System.out.println("Standard Deviation of Column K: " + stdDevK);
        }
    }
}
