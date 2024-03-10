package excelfileoperation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	public static void main(String[] args) {
		

	try {
		String filePath="D:\\Abdhul\\GUVI\\BATCH-JAT16WD\\Writeme.xlsx";
		File file= new File(filePath);
		FileInputStream fis=new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.createSheet("Sheet1");
		XSSFRow Chd = sheet.createRow((short)0);
		Chd.createCell(0).setCellValue("Name"); 
		Chd.createCell(1).setCellValue("Age");  
		Chd.createCell(2).setCellValue("Email");  
		XSSFRow row1 = sheet.createRow((short)1);
		row1.createCell(0).setCellValue("John Doe"); 
		row1.createCell(1).setCellValue("30");  
		row1.createCell(2).setCellValue("john@test.com"); 
		XSSFRow row2 = sheet.createRow((short)2);
		row2.createCell(0).setCellValue("Jane Doe"); 
		row2.createCell(1).setCellValue("28");  
		row2.createCell(2).setCellValue("john@test.com"); 
		XSSFRow row3 = sheet.createRow((short)3);
		row3.createCell(0).setCellValue("Bob Smith"); 
		row3.createCell(1).setCellValue("35");  
		row3.createCell(2).setCellValue("jacky@example.com"); 
		XSSFRow row4 = sheet.createRow((short)4);
		row4.createCell(0).setCellValue("Swapnil"); 
		row4.createCell(1).setCellValue("37");  
		row4.createCell(2).setCellValue("swapnil@example.com"); 
		FileOutputStream fos=new FileOutputStream("D:\\\\Abdhul\\\\GUVI\\\\BATCH-JAT16WD\\\\Writeme.xlsx");
		wb.write(fos);
		System.out.println("Excel Spreedsheet written successfully");
	}catch (Exception e) {
		System.out.println("Exception Occured");
		
	}
}
}
