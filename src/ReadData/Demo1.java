package ReadData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo1 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		String excelpath = ".\\ExcelFiles\\Demo.xlsx";
		
		FileInputStream read = new FileInputStream(excelpath);
		
		String demo = WorkbookFactory.create(read).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
		
		System.out.println(demo);
	}

}
