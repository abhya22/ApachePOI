package ReadData;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel3 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		String path = "E:\\Study\\Excel Files\\Read Excel 3.xlsx";
		
		FileInputStream read = new FileInputStream(path);
		
		double a = WorkbookFactory.create(read).getSheet("Sheet1").getRow(1).getCell(1).getNumericCellValue();
		
		System.out.println(a);
	}

}
