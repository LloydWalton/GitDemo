package basic;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	public static void main(String[] args) throws IOException {
		
		
		
		FileInputStream fis =new FileInputStream("D:\\TestData\\Data.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		
		
		for(int i=0;i<workbook.getNumberOfSheets();i++) {
			
			String sheetName=workbook.getSheetName(i);
			
			if(sheetName.equalsIgnoreCase("Amazon"))
				
			{
				
				XSSFSheet sheet=workbook.getSheetAt(i);
			
				
			}
		}
			
		
		
		
		

	}

}
