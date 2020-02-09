package excelHelpers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {

	static XSSFWorkbook wb;
	static XSSFRow rw;
	static XSSFSheet sheetX;
	static FileOutputStream fileOut;
	static File file;
	static FileInputStream fileIn;

	public static void main(String[] args) {

//		createXLWorkbook("./ComparisonOutput\\workbook.xlsx", "sheet1", 10);
//		openXLWorkbook("./ComparisonOutput\\workbook.xlsx");
//		setXLCellValue("sheet1", 0, 0, "Hello");
//		setXLCellValue("sheet1", 0, 1, "How");
//		setXLCellValue("sheet1", 1, 0, "Are");
//		setXLCellValue("sheet1", 1, 0, "You");
//		writeAndCloseXL("./ComparisonOutput\\workbook.xlsx");
//
//		openXLWorkbook("./ComparisonOutput\\workbook.xlsx");
//		int rowSize = getXLSheetRowCount("Sheet1");
//		System.out.println(rowSize);
//		for (int rowNo = 0; rowNo <= rowSize; rowNo++) {
//			try {
//				System.out.println(getXLSheetData("sheet1", rowNo, 0));
//			} catch (NullPointerException e) {
//				System.out.println("Data is null in row " + rowNo);
//			}
//		}
//		closeXL();
		
		openXLWorkbook("./ComparisonOutput\\workbook.xlsx");
		deleteXLSheet(1);
		closeXL();
	}

	public static void createXLWorkbook(String xlFilePath, String sheetName, int rowSize) {

		wb = new XSSFWorkbook();

		try {
			fileOut = new FileOutputStream(xlFilePath);
			createXLSheet(sheetName, rowSize);

			wb.write(fileOut);
			wb.close();
			fileOut.flush();
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

		System.out.println("Excel workbook \"" + xlFilePath + "\" created successfully");

	}

	// Create a new excel sheet
	public static void createXLSheet(String sheetName, int rowSize) {
		sheetX = wb.createSheet(sheetName);
		for (int r = 0; r < rowSize; r++) {
			sheetX.createRow(r);
		}
		System.out.println("Excel sheet \"" + sheetName + "\" created successfully");
	}

	// Open excel in FileInputStream
	public static void openXLWorkbook(String xlFilePath) {

		file = new File(xlFilePath);

		try {
			fileIn = new FileInputStream(file);
			wb = new XSSFWorkbook(fileIn);
		} catch (Exception e) {
			e.printStackTrace();
		}

		if (file.isFile() && file.exists()) {
			System.out.println("Excel file opened successfully.");
		} else {
			System.out.println("Error in opening excel file.");
		}
	}

	// Set cell value
	public static void setXLCellValue(String sheetName, int rowNum, int colNum, String value) {
		sheetX = wb.getSheet(sheetName);
		sheetX.getRow(rowNum).createCell(colNum).setCellValue(value);
	}

	public static String getXLSheetData(String sheetName, int rowNum, int columnNum) {
		sheetX = wb.getSheet(sheetName);
		String data = sheetX.getRow(rowNum).getCell(columnNum).getStringCellValue();
		return data;
	}

	public static int getXLSheetRowCount(String sheetName) {
		sheetX = wb.getSheet(sheetName);
		int rowSize = sheetX.getLastRowNum();
		return rowSize;
	}

	public static void writeAndCloseXL(String xlFilePath) {
		try {
			fileOut = new FileOutputStream(xlFilePath);
			wb.write(fileOut);
			System.out.println("Data written in excel file successfully");
			wb.close();
			fileOut.flush();
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public static void deleteXLSheet(int  sheetNum) {
		 wb.removeSheetAt(sheetNum);
	}

	public static void closeXL() {
		try {
			wb.close();
			fileIn.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
