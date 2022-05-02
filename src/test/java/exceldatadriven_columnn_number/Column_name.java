package exceldatadriven_columnn_number;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Column_name {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		FileInputStream fis= new FileInputStream("C:\\Users\\home\\Desktop\\TestData.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet("Login");
		XSSFRow row=sheet.getRow(0);
		XSSFCell cell;
		
		int cellindex=1;
		
		for(int i=0;i<row.getLastCellNum();i++)
		{
			if(row.getCell(i).getStringCellValue().trim().equals("password"))
			cellindex=i;
		}
		
		row=sheet.getRow(3);
		cell=row.getCell(cellindex);
		String str=cell.getStringCellValue();
		System.out.println(str);
		
		
		wb.close();
		fis.close();

	}

}
