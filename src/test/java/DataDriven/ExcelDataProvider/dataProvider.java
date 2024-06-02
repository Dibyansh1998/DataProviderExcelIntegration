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

	@Test(dataProvider = "driveTest")
	public void testCaseData(String greeting, String communication, String id) {
		System.out.println(greeting + "||" + communication + "||" + id);

	}

	@DataProvider(name = "driveTest")
	public Object[][] getData() throws IOException {
		String path = "C:\\Users\\dibya\\OneDrive\\Documents\\OneNote Notebooks\\DataDriven_ExcelIntregation.xlsx";
		

		try (FileInputStream fis = new FileInputStream(path); 
				
				XSSFWorkbook wb = new XSSFWorkbook(fis)) 
		{

			XSSFSheet sheet = wb.getSheetAt(0);
			int rowCount = sheet.getPhysicalNumberOfRows();
			int colCount = sheet.getRow(0).getLastCellNum();

			// Create a 2D array to store the data
			Object[][] data = new Object[rowCount - 1][colCount];

			// Iterate through rows and columns to read data
			for (int i = 1; i < rowCount; i++) { // Start from 1 to skip header row
				XSSFRow row = sheet.getRow(i+1);
				for (int j = 0; j < colCount; j++) {
					XSSFCell cell = row.getCell(j);
					data[i][j] = formatter.formatCellValue(cell);
				}
			}
			return data;
		}

		
	}

}