package ExcelData;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.MessageFormat;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class testNGDataProvider {

	DataFormatter dataFormatter = new DataFormatter();

	@Test(dataProvider = "dataProvider")
	public void testCaseData(String testcase, String greeting, String communication, String id) {
		System.out.println(
				MessageFormat.format("Test Data to use: {0}, {1}, {2}, {3}", testcase, greeting, communication, id));
	}

	@DataProvider(name = "dataProvider")
	public Object[][] getData() throws IOException {

		// Locate excel file
		FileInputStream excelDataFileInputStream = new FileInputStream(
				System.getProperty("user.dir") + "//ExcelDataTestNG.xlsx");

		// Read in data
		XSSFWorkbook excelWorkbook = new XSSFWorkbook(excelDataFileInputStream);

		// Find first sheet in excel workbook
		XSSFSheet excelSheet = excelWorkbook.getSheetAt(0);

		// Identify how many populated rows the excel sheet has
		int rowCount = excelSheet.getPhysicalNumberOfRows();

		// Retrieve the first row in the excel sheet
		XSSFRow excelRow = excelSheet.getRow(0);

		// Determine how many columns in row
		int columnCount = excelRow.getLastCellNum();

		Object testDataObjects[][] = new Object[rowCount - 1][columnCount];

		// Iterate through rows
		for (int rows = 0; rows < rowCount - 1; rows++) {

			// Find the first line of test data in excel sheet (ignore header)
			excelRow = excelSheet.getRow(rows + 1);

			// Iterate through columns
			for (int columns = 0; columns < columnCount; columns++) {
				if (excelRow.getCell(columns) != null) {
					XSSFCell excelCell = excelRow.getCell(columns);

					// Use the DataFormatter to convert to correct type
					testDataObjects[rows][columns] = dataFormatter.formatCellValue(excelCell);
				}
			}
		}
		return testDataObjects;
	}
}
