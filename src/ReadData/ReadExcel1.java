package ReadData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel1 {

	public static void main(String[] args) throws IOException {
		
		String excelpath = ".\\ExcelFiles\\City.xlsx";            //For Locating file path
		
		FileInputStream read = new FileInputStream(excelpath);    //For Read Data which will present in XLXS
		
//		* Step By Step activities *
		
//		1. Open file from stream mode
//		2. Get the workbook from that file
//		3. Get the Sheet
//		4. Get the rows
//		5. Finally read the data from the cell
		
		XSSFWorkbook wb = new XSSFWorkbook(read);
		
		XSSFSheet sheet = wb.getSheet("Sheet1");             //This Method will get the sheet (Specify name of the sheet), this return sheet object.

		// If you want to pass index there is one more method
		
		// XSSFSheet sheet = wb.getSheetAt(0);               // This Method will pass index
		
		//Using For Loop For Read All the data from xlx file
		
			int rows = sheet.getLastRowNum();
			int column = sheet.getRow(rows).getLastCellNum();
			
			for(int r=0; r<=column; r++) {                    // First For Loop is for Rows
				
				XSSFRow row = sheet.getRow(r);
				
				for(int c=0; c<column; c++) {                 // First For Loop is for column
					
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType()) {
					
				case STRING: System.out.print(cell.getStringCellValue()); 
				break;
				case NUMERIC: System.out.print(cell.getStringCellValue());
				break;
				case BOOLEAN: System.out.print(cell.getStringCellValue());
				
				}
				System.out.print("  ");
				}
				System.out.println();
			}
	}
}
