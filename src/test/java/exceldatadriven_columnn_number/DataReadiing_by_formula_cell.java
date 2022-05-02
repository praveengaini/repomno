package exceldatadriven_columnn_number;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataReadiing_by_formula_cell {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		String Path= "C:\\Users\\home\\Desktop\\TestData.xlsx";
		FileInputStream fis =new FileInputStream(Path);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(1);
		
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
				
				
				System.out.print("|");				
			}
		
			System.out.println();
		}
		
		wb.close();
		fis.close();

	}

}
