package ExcelData;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataExtractionXSSFWorkbook {

	public static void main(String[] args) throws IOException {
		getData("Purchase");
	}

	public static ArrayList<String> getData(String testCaseName) throws IOException {
		// Test data array
		ArrayList<String> testDataArrayList = new ArrayList<String>();

		// Read in the file
		FileInputStream fileInputStream = new FileInputStream(
				System.getProperty("user.dir") + "//ExcelDataXSSFWorkbook.xlsx");

		// Add utility to manipulate workbook
		XSSFWorkbook excelWorkbook = new XSSFWorkbook(fileInputStream);

		// Find all Excel sheets
		int numSheets = excelWorkbook.getNumberOfSheets();

		for (int i = 0; i < numSheets; i++) {
			if (excelWorkbook.getSheetName(i).equalsIgnoreCase("ExcelSheet01")) {
				// Find the desired Excel sheet
				XSSFSheet excelSheet = excelWorkbook.getSheetAt(i);

				// Read all the rows of the excel sheet
				Iterator<Row> excelRowIterator = excelSheet.rowIterator();

				// Read the excel sheet header row
				Row excelRow = excelRowIterator.next();

				// Scan the columns to identify correct column
				Iterator<Cell> rowCells = excelRow.cellIterator();

				int counter = 0;
				int column = 0;

				while (rowCells.hasNext()) {
					Cell cell = rowCells.next();

					// Find exact column index
					if (cell.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column = counter;
					}
					counter++;
				}
				System.out.println(MessageFormat.format("TestCases column index: {0}", column));

				// Scan column values to identify test case row
				while (excelRowIterator.hasNext()) {
					Row testCaseRow = excelRowIterator.next();
					if (testCaseRow.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {

						// Retrieve test case data from row
						Iterator<Cell> testCaseCellsIterator = testCaseRow.cellIterator();
						while (testCaseCellsIterator.hasNext()) {

							Cell testDataCell = testCaseCellsIterator.next();

							// Action to take if cell value is a string
							if (testDataCell.getCellType() == CellType.STRING) {
								// Print each cell data out to console
								String cellTestDataString = testDataCell.getStringCellValue();
								System.out.println(MessageFormat.format("Cell test data: {0}", cellTestDataString));
								testDataArrayList.add(cellTestDataString);
							} else {

								// Convert a numerical value to a string
								testDataArrayList.add(NumberToTextConverter.toText(testDataCell.getNumericCellValue()));
								System.out.println(MessageFormat.format("Cell test data: {0}",
										testDataCell.getNumericCellValue()));
							}

						}
					}
				}
			}

		}
		return testDataArrayList;
	}

}
