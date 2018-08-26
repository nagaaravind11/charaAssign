import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ExcelWriter {

   
    public static void addToExcel(List<String> colNames, List<List<String>> colValuesList , String sheetname ,String excelname) throws IOException {
        
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook();    

      
        // Create a Sheet
        Sheet sheet = workbook.createSheet(sheetname);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Creating cells
        for(int i = 0; i < colNames.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(colNames.get(i));
            cell.setCellStyle(headerCellStyle);
        }


        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(List<String> colValueList: colValuesList) {
            Row row = sheet.createRow(rowNum++);
            for(int i=0 ;i<colValueList.size() ;i++)
            {
                row.createCell(i)
                .setCellValue(colValueList.get(i));
            }
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < colNames.size(); i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(excelname);
        workbook.write(fileOut);
        fileOut.close();

        workbook.close();
    }


  
    
}

