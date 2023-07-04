package ExcelData;

import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;

public class testExample {

	public static void main(String[] args) throws IOException {

		dataExtractionXSSFWorkbook extractedDataExtraction = new dataExtractionXSSFWorkbook();
		ArrayList<String> testDataArrayList = extractedDataExtraction.getData("Add Profile");

		for (String testDataCell : testDataArrayList) {
			System.out.println(MessageFormat.format("Test data extracted: {0}", testDataCell));
		}

	}

}
