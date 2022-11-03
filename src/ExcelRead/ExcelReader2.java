package ExcelRead;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
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

public class ExcelReader2{

	public static void main(String[] args) throws IOException {

		
		FileInputStream fs = new FileInputStream("D:\\DemoFile.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fs);
		XSSFSheet sheet = wb.getSheetAt(0);
		
		Row rows = sheet.getRow(0);
		Cell cell = rows.getCell(0);
		
		System.out.println(sheet.getRow(0).getCell(0));
		
		Row rows1 = sheet.getRow(1);
		Cell cell1 = rows1.getCell(1);
		
		System.out.println(sheet.getRow(0).getCell(1));
		
		Row rows2 = sheet.getRow(2);
		Cell cell2 = rows2.getCell(2);
		
		System.out.println(sheet.getRow(1).getCell(2));
		
		DataFormatter formatter = new DataFormatter();
		
		Iterator<Row>it = sheet.rowIterator();
		while(it.hasNext()) {
			Row row = it.next();
		
		System.out.println(formatter.formatCellValue(rows.getCell(0)));
		System.out.println(formatter.formatCellValue(rows.getCell(1)));

		System.out.println(formatter.formatCellValue(rows.getCell(2)));

		System.out.println(formatter.formatCellValue(rows.getCell(3)));

		System.out.println(formatter.formatCellValue(rows.getCell(4)));
		}
		
		
//		XSSFWorkbook wb1 = new XSSFWorkbook(fs);
//		XSSFSheet sheet1 = wb.createSheet("Sheet2");
//		Row row = sheet.createRow(0);
//		row.createCell(0).setCellValue(null);
	
	}

}
