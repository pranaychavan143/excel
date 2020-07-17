package com.excelutil.util;
import com.excelutil.model.Invoice;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

public class CreateInvoices {

    public static void main(String []args){
        try {
            //Create Workbook in .xlsx formate
                Workbook workbook = new XSSFWorkbook();
                //Create Sheet
            Sheet sheet =workbook.createSheet();
            String[] columnHeadings ={"Item id", "Item Name", "Qty", "Item Price", "Sold Date"};
            //Make it bold with foreground color
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short)12);
            headerFont.setColor(IndexedColors.BLACK.index);
            //Cell Style with Font
            CellStyle hederStyle = workbook.createCellStyle();
            hederStyle.setFont(headerFont);
            hederStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            hederStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
            //Create The Header Row
            Row hedeRow = sheet.createRow(0);
            //Iterate Over The column headings  to create columns
            for(int  i=0;i<columnHeadings.length;i++){
                Cell cell= hedeRow.createCell(i);
                cell.setCellValue(columnHeadings[i]);
                cell.setCellStyle(hederStyle);
            }
            //Fill The Data

            ArrayList<Invoice> a = createData();
            CreationHelper creationHelper = workbook.getCreationHelper();
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("mm/dd/yyyy"));
            int rownum =1;
            for(Invoice i : a) {

                Row row = sheet.createRow(rownum++);
                row.createCell(0).setCellValue(i.getItemId());
                row.createCell(1).setCellValue(i.getItemName());
                row.createCell(2).setCellValue(i.getItemQty());
                row.createCell(3).setCellValue(i.getTotalPrice());
                Cell dateCell = row.createCell(4);
                dateCell.setCellValue(i.getItemSoldDate());
                dateCell.setCellStyle(dateStyle);
            }
            //AutoSize column
            for(int i=0; i<columnHeadings.length;i++){
                sheet.autoSizeColumn(i);
            }

            Sheet sheet1 = workbook.createSheet("Second");
            //Write the output file

            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Pranay\\Downloads\\Invoice.xlxs");
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
            System.out.println("Completed");
 
        }
        catch (Exception e){
            e.printStackTrace();
        }
    }
    private static  ArrayList<Invoice>createData() throws ParseException
    {
        ArrayList<Invoice> a = new ArrayList();
        a.add(new Invoice(1, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(2, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(3, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(4, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(5, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(6, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(7, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(8, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(9, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(10, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(11, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(12, "Pen", 100, 20.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(13, "Book", 2, 10.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        a.add(new Invoice(14, "Table", 1, 50.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/02/2020")));
        a.add(new Invoice(15, "Lamp", 5, 100.0, new SimpleDateFormat("MM/dd/yyyy").parse("01/01/2020")));
        return a;
    }
}
