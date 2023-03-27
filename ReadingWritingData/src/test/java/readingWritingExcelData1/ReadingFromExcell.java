package readingWritingExcelData1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingFromExcell {

	public static void main(String[] args) throws IOException {
		/*
		 * String path = "C:\\ReadingTest\\Microsoft Excel  xlsx.xlsx";
		 * 
		 * File src = new File(path);
		 * 
		 * FileInputStream fis = new FileInputStream(src);
		 * 
		 * XSSFWorkbook workbook = new XSSFWorkbook(fis);
		 * 
		 * XSSFSheet sheet = workbook.getSheetAt(0);
		 * 
		 * for (Row row : sheet) { for (Cell cell : row) {
		 * 
		 * System.out.println(cell.getStringCellValue() + "\t"); } System.out.println();
		 * }
		 * 
		 * workbook.close(); fis.close();
		 */

		String path = "C:\\ExcelData\\New Microsoft Excel Worksheet.xlsx.xlsx";

		File src = new File(path);

		FileInputStream fis = new FileInputStream(src);

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		XSSFSheet sheet = workbook.getSheetAt(0);

		int rows = sheet.getLastRowNum();  //return number of rows
		int cols = sheet.getRow(1).getLastCellNum();//return number of cell in particular rows

		/*
		 * for (Row row : sheet) {
		 *  for (Cell cell : row) {
		 * System.out.println(cell.getStringCellValue() + "\t"); } System.out.println();
		 * }
		 */

		for (int r = 0; r <= rows; r++) {

			XSSFRow row = sheet.getRow(r);

			for (int c = 0; c < cols; c++) {

				XSSFCell cell = row.getCell(c);//cell from particular row

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;

				}
				System.out.print("\t");
			}
			System.out.println();
			workbook.close();

			fis.close();
		}
	}

}

