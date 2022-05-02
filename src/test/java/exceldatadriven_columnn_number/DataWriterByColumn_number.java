package exceldatadriven_columnn_number;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataWriterByColumn_number {
	
	
	public static void main(String[] args) throws Exception {
		
		
		FileInputStream fis =new FileInputStream("C:\\Users\\home\\Desktop\\TestData.xlsx");
		FileOutputStream fos= null;
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Login");
		XSSFRow row = sheet.getRow(2);
		XSSFCell cell = row.getCell(2);
		
		cell.setCellValue("jkm");
		fos= new FileOutputStream("C:\\Users\\home\\Desktop\\TestData.xlsx");
		wb.write(fos);
		wb.close();
		fos.close();
		System.out.println("write opeartion succesfully completed");
		
		
		
	
		
		
		
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}


	

	
		
		
	
	
	
	
	
	
	
	
