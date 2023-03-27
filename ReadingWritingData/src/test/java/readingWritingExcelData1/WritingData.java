package readingWritingExcelData1;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingData {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("Employees");

		Object empdata[][] = { { "EmpId", "FName","LName", "Job", "DeptId","PhoneNumber","HireDate","Salary","Bonus","Commission","Sex","BirthDate" },

				{ 101, "Mr","David" ,"Manager",30,"3474532273","8-25-22",150000,"20 pct","20 pct","Male","1-23-1-2000" },
				{ 102, "Saiful","Chowdhury" ,"PM" ,20,"3470536233","6-27-2020",110000,"10 pct","20 pct","Male","1-23-1-1990" },
				{ 103, "Nuyamah", "Rukayat","BA" ,30,"34745399233","1-12-2019",120000,"20 pct","20 pct","Female","1-23-1-1995"},
				{ 104, "Nabeeha","Rukayath","Pm" ,30,"3474508233","12-11-2018",100000,"10 pct","10 pct","Female","1-23-1-2002"},
                { 105, "Sultan","Chowdhury","Pm" ,40,"3474537233","6-4-2023",80000,"20 pct","10 pct","Male","1-23-1-1999"},
				{ 106, "Nazmul","Islam","Scrum Master" ,20,"3474235233","10-12-2017",1500000,"5 pct","10 pct","Male","1-23-1-1994"},
				{ 107, "Nurul","Islam","Po" ,50,"3474536243","6-19-2018",110000,"5 pct","10 pct","Male","1-23-1-1995"},
				{ 108, "Sabina","Yeasmin","po" ,50,"3474546213","7-12-2022",130000,"10 pct","5 pct","Female","1-23-1-2001"},
				{ 109, "Wahidur","Rahman","Developer" ,20,"3473536233","6-22-2021",140000,"10 pct","10 pct","Male","1-23-1-1998"},
				{110, "Tahmina","Akter","Developer" ,30,"3474536203","4-12-2014",155000,"20 pct","5 pct","Female","1-23-1-1996"},
				{ 204, "Nasrin","Sultana","Pm" ,40,"3474536439","6-22-2023",110000,"10 pct","10 pct","Female","1-23-1-2003"},
				{ 304, "Sakhaoat","Hossain","Enginner" ,20,"3474576233","6-10-2012",160000,"5 pct","20 pct","Male","1-23-1-1990"},
				
				
				
				
				
				
				
				
		};
		
		
		//using for loop

		int rows = empdata.length;
		int cols = empdata[0].length;

		System.out.println(rows);

		System.out.println(cols);

		for (int r = 0; r < rows; r++) {

			XSSFRow row = sheet.createRow(r);
			for (int c = 0; c < cols; c++) {

				XSSFCell cell = row.createCell(c);

				Object value = empdata[r][c];

				if (value instanceof String)

					cell.setCellValue((String) value);

				if (value instanceof Integer)

					cell.setCellValue((Integer) value);

				if (value instanceof Boolean)

					cell.setCellValue((Boolean) value);

			}

		}

		String filePath = "C:\\Test_Data\\excell.xlsx.xlsx";

		FileOutputStream outStream = new FileOutputStream(filePath);
		workbook.write(outStream);

		outStream.close();

		System.out.println("Employee xls file created successfully");

	}

}

