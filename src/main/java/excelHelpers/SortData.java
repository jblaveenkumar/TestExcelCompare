package excelHelpers;

import com.aspose.cells.CellArea;
import com.aspose.cells.Cells;
import com.aspose.cells.DataSorter;
import com.aspose.cells.SortOrder;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class SortData {
	public static void main(String[] args) throws Exception {

		String dataDir = "D:\\Eclipse Photon\\Workspace\\TestExcelCompare\\ComparisonOutput\\";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Workbook.xlsx");
		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		// Get the cells collection in the sheet
		Cells cells = worksheet.getCells();

		// Obtain the DataSorter object in the workbook
		DataSorter sorter = workbook.getDataSorter();
		// Set the first order
		sorter.setOrder1(SortOrder.ASCENDING);
		// Define the first key.
		sorter.setKey1(1);
//		// Set the second order
//		sorter.setOrder2(SortOrder.ASCENDING);
//		// Define the second key
//		sorter.setKey2(1);

		// Create a cells area (range).
		CellArea ca = new CellArea();
		// Specify the start row index.
		ca.StartRow = 0;
		// Specify the start column index.
		ca.StartColumn = 0;
		// Specify the last row index.
		ca.EndRow = 200;
		// Specify the last column index.
		ca.EndColumn = 4;
		// Sort data in the specified data range (A2:C10)
		sorter.sort(cells, ca);
//		workbook.getWorksheets().removeAt("Sheet2");
		// Saving the excel file
		workbook.save(dataDir + "Workbook2.xlsx");
		ExcelUtility.openXLWorkbook(dataDir + "Workbook2.xlsx");
		ExcelUtility.deleteXLSheet(1);
		ExcelUtility.writeAndCloseXL(dataDir + "Workbook2.xlsx");
	}
}
