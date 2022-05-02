package exceldatadriven_columnn_number;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataReading_pwd_protected_file {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
	
		String path ="C:\\Users\\home\\Desktop\\pwdexample.xlsx";
		FileInputStream fis =new FileInputStream(path);
		String pwd ="Test@123";
		
		 //Workbook wb=WorkbookFactory.create(path);
		
		XSSFWorkbook wb = (XSSFWorkbook) WorkbookFactory.create(fis, pwd);
		 XSSFSheet sheet =wb.getSheetAt(0);
			
			int rows=sheet.getLastRowNum();
			int cols=sheet.getRow(0).getLastCellNum();
			
			for(int r=0;r<=rows;r++)
			{
				XSSFRow row=sheet.getRow(r);
				
				for(int c=0;c<cols;c++)
				{
					XSSFCell cell=row.getCell(c);
					switch(cell.getCellType())
					{
					
						case STRING:  System.out.print(cell.getStringCellValue());break;
						case NUMERIC: System.out.print(cell.getNumericCellValue());break;
						case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
						case FORMULA: System.out.print(cell.getNumericCellValue());break;
					}
					
					
					System.out.print(" | ");				
				}
			
				System.out.println();
		
			}
		

	}

}
