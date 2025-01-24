

package excelReadandWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class B7ExcelReadWriteCls {
	public static void main(String[] args) {
		B7ExcelReadWriteCls x=new B7ExcelReadWriteCls();
		
		for (int i=0;i<4;i++) {  // i - row
			for (int j=0;j<4;j++) {  //j - column
				System.out.print(x.getExcelData("Sheet1",i, j)+ "   ");
			}
			System.out.println( " ");
		}
		
		x.WriteToExcel("Sheet1", 0, 3, "Result");
		x.WriteToExcel("Sheet1", 1, 3, "Fail");
		x.WriteToExcel("Sheet1", 2, 3, "Pass");
		x.WriteToExcel("Sheet1", 3, 3, "Pass");
						
	}
	public void WriteToExcel(String Sheetname, int rowNum, int cellNum, String desc) {
		FileInputStream fis;
		XSSFWorkbook wb;
		
		try {
			fis =new FileInputStream("Util//Employee.xlsx");
			wb=new XSSFWorkbook(fis);
			XSSFSheet s = wb.getSheet(Sheetname);
			XSSFRow r = s.getRow(rowNum);
			XSSFCell c=r.createCell(cellNum);
			c.setCellValue(desc);
			FileOutputStream fos=new FileOutputStream("Util//Employee.xlsx");
			wb.write(fos);
			
			} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	
		
	}
				
		
	
	public String getExcelData(String sheetName, int rowNum, int colNum) {
		String retVal = null;
		try {
			FileInputStream fis = new FileInputStream("Util//Employee.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet s = wb.getSheet(sheetName);
			XSSFRow r = s.getRow(rowNum);
			XSSFCell c = r.getCell(colNum);
			retVal = B7ExcelReadWriteCls.getCellValue(c);
			fis.close();
			wb.close();

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return retVal;
	}

	public static String getCellValue(XSSFCell c) {
		switch(c.getCellType()) {
		case NUMERIC:
			return String.valueOf(c.getNumericCellValue());  //10 -> "10"
		case BOOLEAN:
			return String.valueOf(c.getBooleanCellValue());
		case STRING:
			return c.getStringCellValue();
		default:
			return c.getStringCellValue();
			
		}
	

	}

}
