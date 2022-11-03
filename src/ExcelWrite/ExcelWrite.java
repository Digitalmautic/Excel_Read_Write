package ExcelWrite;

import java.io.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
  
public class ExcelWrite {
	
	public static void main(String[] args) {
  
	    XSSFWorkbook workbook = new XSSFWorkbook();
	    XSSFSheet sheet = workbook.createSheet("Calculate Simple Interest");
	  
	    Row header = sheet.createRow(0);
	    header.createCell(0).setCellValue("Pricipal");
	    header.createCell(1).setCellValue("RoI");
	    header.createCell(2).setCellValue("T");
	    header.createCell(3).setCellValue("Interest (P r t)");
	      
	    Row dataRow = sheet.createRow(1);
	    dataRow.createCell(0).setCellValue(14500d);
	    dataRow.createCell(1).setCellValue(9.25);
	    dataRow.createCell(2).setCellValue(3d);
	    dataRow.createCell(3).setCellFormula("A2*B2*C2");
	      
	    try {
	        FileOutputStream out =  new FileOutputStream(new File("D:\\xformulaDemo.xlsx"));
	        workbook.write(out);
	        out.close();
	        System.out.println("Excel with foumula cells written successfully"); }
	          
	     catch (FileNotFoundException e) 
	    {
	        e.printStackTrace(); }
	    
	     catch (IOException e) 
	    {
	        e.printStackTrace(); }
    
}
}