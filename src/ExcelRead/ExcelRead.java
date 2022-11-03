package ExcelRead;

import java.io.File;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
public static void main (String [] args) throws IOException{
	
	
FileInputStream fs = new FileInputStream("D:\\DemoFile.xlsx");

XSSFWorkbook workbook = new XSSFWorkbook(fs);
XSSFSheet sheet = workbook.getSheetAt(0);

Row row = sheet.getRow(0);
Cell cell = row.getCell(0);
System.out.println(sheet.getRow(0).getCell(0));  // 662.0

Row row1 = sheet.getRow(1);
Cell cell1 = row1.getCell(1);
System.out.println(sheet.getRow(0).getCell(1));  // Nausira

Row row2 = sheet.getRow(1);
Cell cell2 = row2.getCell(1);
System.out.println(sheet.getRow(1).getCell(2));  // 663.0

Row row3 = sheet.getRow(1);
Cell cell3 = row3.getCell(1);
System.out.println(sheet.getRow(2).getCell(3));  // Baripada


// DataFormatter contains methods for formatting the value stored in an Cell. 
DataFormatter formatter = new DataFormatter();

Iterator<Row> rowIterator = sheet.iterator();

while (rowIterator.hasNext()) {

    Row rows = rowIterator.next();
    
    // Since I needed only five columns, I hardcoded the column numbers
    
    System.out.println(formatter.formatCellValue(rows.getCell(0)));
    System.out.println(formatter.formatCellValue(rows.getCell(1)));
    System.out.println(formatter.formatCellValue(rows.getCell(2)));
    System.out.println(formatter.formatCellValue(rows.getCell(3)));
    System.out.println(formatter.formatCellValue(rows.getCell(4)));
    System.out.println(formatter.formatCellValue(rows.getCell(5)));      	
}
}
}