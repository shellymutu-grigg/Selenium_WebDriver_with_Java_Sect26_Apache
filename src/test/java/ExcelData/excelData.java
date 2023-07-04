package ExcelData;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.MessageFormat;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class excelData {

	@Test
	public void getExcelData() throws IOException {

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
			System.out.println(MessageFormat.format("Outer row loop {0}", rows));
			// Find the first line of test data in excel sheet (ignore header)
			excelRow = excelSheet.getRow(rows + 1);

			// Iterate through columns
			for (int columns = 0; columns < columnCount; columns++) {
				System.out.println(MessageFormat.format("Inner column loop {0}", columns));
				if (excelRow.getCell(columns) != null) {
					XSSFCell excelCell = excelRow.getCell(columns);

					// Use the DataFormatter to convert to correct type
					System.out.println(MessageFormat.format("Cell value: {0}", excelRow.getCell(columns)));
				}
			}
		}
	}

}
