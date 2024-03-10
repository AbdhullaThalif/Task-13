package excelfileoperation;


import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	public static void main(String[] args) {
		
		ReadExcel objRE=new ReadExcel();
		String StudentDt=objRE.ReadExcel();
		System.out.println(StudentDt);
		
		}

		public String ReadExcel() {
			// TODO Auto-generated method stub
			String data="";
			
			try {
				XSSFWorkbook wb=new XSSFWorkbook("D:\\\\Abdhul\\\\GUVI\\\\BATCH-JAT16WD\\\\Readme.xlsx");
				XSSFSheet sheet=wb.getSheet("Sheet1");
				
				XSSFRow row=sheet.getRow(0);
				data=row.getCell(0).getStringCellValue();
				
			}catch (Exception e) {
				System.out.println("Reading excel code has some problems");
				e.printStackTrace();
			}
			return data;	
			}

}
