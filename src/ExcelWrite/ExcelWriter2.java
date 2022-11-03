package ExcelWrite;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter2 {

	public static void main(String[] args) {


		 XSSFWorkbook workbook = new XSSFWorkbook();
		    XSSFSheet sheet = workbook.createSheet("Calculate Simple Interest");
		  
		    Row row1 = sheet.createRow(0);
		    row1.createCell(0).setCellValue("Pricipal");
		    row1.createCell(1).setCellValue("RoI");
		    row1.createCell(2).setCellValue("is");
		    row1.createCell(3).setCellValue("5000000");
		      
		    Row row2 = sheet.createRow(1);
		    row2.createCell(0).setCellValue("14500d");
		    row2.createCell(1).setCellValue("9.25");
		    row2.createCell(2).setCellValue("3d");
		    row2.createCell(3).setCellFormula("%66778");
		      
		    try {
		        FileOutputStream out =  new FileOutputStream(new File("D:\\DemoExcelReader2.xlsx"));
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
