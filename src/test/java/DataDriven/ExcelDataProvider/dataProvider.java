package DataDriven.ExcelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvider {

	DataFormatter formatter = new DataFormatter();

	@Test(dataProvider = "driveTest", groups="Sanity")
	public void testCaseData(String greeting, String communication, String id) {
		System.out.println(greeting + "||" + communication + "||" + id);

	}

	@DataProvider(name = "driveTest")
	public Object[][] getData() throws IOException {
		String path = "C:\\Desktop Items\\Excel\\ExcelDataProvider.xlsx";
		

		try (FileInputStream fis = new FileInputStream(path); 
				
				XSSFWorkbook wb = new XSSFWorkbook(fis)) 
		{

			XSSFSheet sheet = wb.getSheetAt(0);
			int rowCount = sheet.getPhysicalNumberOfRows();
			XSSFRow row=sheet.getRow(0);
			int colCount = row.getLastCellNum();

			// Create a 2D array to store the data
			Object[][] data = new Object[rowCount-1][colCount];

			// Iterate through rows and columns to read data
			for (int i = 0; i < rowCount-1; i++) 
			{ // Start from 1 to skip header row
					 row=sheet.getRow(i+1);
					 
				for (int j = 0; j < colCount; j++) 
				{
					XSSFCell cell = row.getCell(j);
					
					data[i][j] = formatter.formatCellValue(cell);
				}
			}
			return data;
		}

		
	}

}