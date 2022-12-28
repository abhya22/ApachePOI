package ReadData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo2 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		String path = "C:\\Users\\SONY\\Desktop\\ExcelFiles\\Demo1.xlsx";
		
		FileInputStream read = new FileInputStream(path);
		
		XSSFWorkbook wb = new XSSFWorkbook(read);
		
		XSSFSheet read1 = wb.getSheet("Sheet1");
		
		int rows = read1.getLastRowNum();
		int column = read1.getRow(rows).getLastCellNum();
		
		for(int r=0; r<=column; r++) {                    // First For Loop is for Rows
			
			XSSFRow row = read1.getRow(r);
			
			for(int c=0; c<column; c++) {                 // First For Loop is for column
				
			XSSFCell cell = row.getCell(c);
			
			switch(cell.getCellType()) {
				
			case STRING: System.out.print(cell.getStringCellValue()); 
			break;
			case NUMERIC: System.out.print(cell.getStringCellValue());
			break;
			case BOOLEAN: System.out.print(cell.getStringCellValue());
			
			}
			System.out.print(" | ");
			}
			System.out.println();
		
	}

}
}
