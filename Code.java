import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Code {

    // Date format for formatting date values
    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("MM-dd-yyyy hh:mm a");

    public static void main(String[] args) {
        // Specify the path to your Excel file
        String filePath = "C:/Users/91885/Downloads/Assignment_Timecard.xlsx";

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                // Retrieve values from the Excel sheet
                String positionId = getCellValueAsString(row.getCell(0));
                String positionStatus = getCellValueAsString(row.getCell(1));
                Date timeIn = getCellValueAsDate(row.getCell(2));
                Date timeOut = getCellValueAsDate(row.getCell(3));
                double timecardHours = getCellValueAsNumeric(row.getCell(4));
                Date payCycleStartDate = getCellValueAsDate(row.getCell(5));
                Date payCycleEndDate = getCellValueAsDate(row.getCell(6));
                String employeeName = getCellValueAsString(row.getCell(7));
                String fileNumber = getCellValueAsString(row.getCell(8));

                // Calculate time difference and check if dates are consecutive
                double timeDifference = calculateTimeDifference(timeIn, timeOut);
                boolean isConsecutiveDays = areConsecutiveDays(payCycleStartDate, payCycleEndDate);

                // Print the details for each row
                printDetails(positionId, positionStatus, timeIn, timeOut, timecardHours,
                        payCycleStartDate, payCycleEndDate, employeeName, fileNumber,
                        timeDifference, isConsecutiveDays);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Print details of a timecard entry
    private static void printDetails(String positionId, String positionStatus, Date timeIn, Date timeOut,
                                     double timecardHours, Date payCycleStartDate, Date payCycleEndDate,
                                     String employeeName, String fileNumber, double timeDifference,
                                     boolean isConsecutiveDays) {
        System.out.println("Position ID: " + positionId);
        System.out.println("Position Status: " + positionStatus);
        System.out.println("Time In: " + formatDate(timeIn));
        System.out.println("Time Out: " + formatDate(timeOut));
        System.out.println("Timecard Hours: " + timecardHours);
        System.out.println("Pay Cycle Start Date: " + formatDate(payCycleStartDate));
        System.out.println("Pay Cycle End Date: " + formatDate(payCycleEndDate));
        System.out.println("Employee Name: " + employeeName);
        System.out.println("File Number: " + fileNumber);
        System.out.println("Time Difference: " + timeDifference);
        System.out.println("Consecutive Days: " + isConsecutiveDays);
        System.out.println("-----------------------------");
    }

    // Format a date to a string or provide "N/A" if the date is null
    private static String formatDate(Date date) {
        return (date != null) ? DATE_FORMAT.format(date) : "N/A";
    }

    // Get cell value as a string or return an empty string if the cell is null
    private static String getCellValueAsString(Cell cell) {
        return (cell != null) ? cell.toString() : "";
    }

    // Get cell value as a date or return null if the cell is not a date
    private static Date getCellValueAsDate(Cell cell) {
        return (cell != null && cell.getCellType() == CellType.NUMERIC) ? cell.getDateCellValue() : null;
    }

    // Get cell value as a numeric or return 0.0 if the cell is not numeric
    private static double getCellValueAsNumeric(Cell cell) {
        return (cell != null && cell.getCellType() == CellType.NUMERIC) ? cell.getNumericCellValue() : 0.0;
    }

    // Calculate time difference between two dates in hours
    private static double calculateTimeDifference(Date start, Date end) {
        return (start != null && end != null) ? Math.abs(end.getTime() - start.getTime()) / (60.0 * 60.0 * 1000) : 0.0;
    }

    // Check if two dates are consecutive (placeholder, replace with actual logic)
    private static boolean areConsecutiveDays(Date startDate, Date endDate) {
        return true;
    }
}
