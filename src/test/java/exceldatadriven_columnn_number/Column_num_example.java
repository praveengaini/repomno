package exceldatadriven_columnn_number;

import java.io.FileInputStream;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Column_num_example {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		FileInputStream fis =new FileInputStream("C:\\Users\\home\\Desktop\\TestData.xlsx");
		XSSFWorkbook wrk=new XSSFWorkbook(fis);
		XSSFSheet sheet=wrk.getSheet("Login");
		XSSFRow row =sheet.getRow(2);
		XSSFCell cell=row.getCell(1);
		 String str=cell.getStringCellValue();
		 System.out.println(str);
		
		

	}

}
