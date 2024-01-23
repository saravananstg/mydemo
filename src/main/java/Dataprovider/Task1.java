package Dataprovider;


import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class Task1 {

	@Test(dataProvider = "dataProviderMethod")
	public void myTestMethod(String param1, int param2) {
	    // Your test logic here
	}

	public void excelWrite() throws Exception {
		String path = "./Data/TestData2.xlsx";
		
		FileInputStream fs = new FileInputStream(path);
		
		Workbook wb = WorkbookFactory.create(fs);
		
		Sheet sheet1 = wb.getSheetAt(0);
		
		int lastRow = sheet1.getLastRowNum();
		for(int i=0; i<=lastRow; i++){
		Row row = sheet1.getRow(i);
		Cell cell = row.createCell(2);

		cell.setCellValue("WriteintoExcel");

		}

		FileOutputStream fos = new FileOutputStream(path);
		wb.write(fos);
		fos.close();
	}
	
}




