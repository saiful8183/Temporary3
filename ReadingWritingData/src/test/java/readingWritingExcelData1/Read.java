package readingWritingExcelData1;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Read {

	public static void main(String[] args) throws Exception, Exception {

		String path = "C:\\ExcelData\\New Microsoft Excel Worksheet.xlsx.xlsx";  //file path

		
		
		File src = new File(path);  // Object of file classh'';l
		
		
		
		//Here I did not create object/instance  of fileinputstream class because i am using "workbookfactory"
		//"workbook factory" has the capability of taking the source file directly.

		
		
		// Creat an instance of workbook class which represents the excell file.
		// This is done by passing a file object

		Workbook workbook = WorkbookFactory.create(src);   // This can be exemple of Polimorphic refference
		
		
		//From workbook object get the sheet which you want to read
		
		Sheet sheet=workbook.getSheetAt(0);
		

		
		//Itorate through rows and columns of sheet
		
		
		
		 for (Row row : sheet) {
		  for (Cell cell : row) {
			  
			  switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue()+"\t");
					break;
//				case NUMERIC:
//					System.out.print(cell.getNumericCellValue()+"\t");
//			  break;
//			  
			  
//			  
//				case BOOLEAN:
//					System.out.print(cell.getBooleanCellValue()+"\t");
//			  break;
//		
			  
			  }
			     //System.out.print(cell.getStringCellValue() + "\t"); } System.out.println();
		 }
		  
			  System.out.println();
		
		
	}
		 workbook.close();
	}
		 
}