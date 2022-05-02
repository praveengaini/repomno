package exceldatadriven_columnn_number;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataReadingfromfile {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		String path="C:\\Users\\home\\Desktop\\TestData.xlsx";
		FileInputStream fis=new FileInputStream(path);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheetAt(0);
		
		int rows=sheet.getLastRowNum();
		System.out.println(rows);
		int cols=sheet.getRow(0).getLastCellNum();
		System.out.println(cols);
		
		
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
				   default:break;
					
				 }
				 
				System.out.print(" | "); 
			}
			
			System.out.println();
		}
		
		wb.close();
		fis.close();
		
		
		
		
		

	}

}
