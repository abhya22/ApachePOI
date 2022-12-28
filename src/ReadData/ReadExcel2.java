package ReadData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel2 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		String path = "C:\\Users\\SONY\\Desktop\\ExcelFiles\\Wah.xlsx";
		
		FileInputStream read = new FileInputStream(path);
		
		//String a = WorkbookFactory.create(read).getSheet("Sheet1").getRow(1).getCell(4).getStringCellValue();
		double a = WorkbookFactory.create(read).getSheet("Sheet1").getRow(1).getCell(4).getNumericCellValue();
		
		System.out.println(a);
		
	}
}

// wah kya baat hai 1
//  0   1    2   3  4


// mum mum zal ka 2
//  0   1   2  3  4