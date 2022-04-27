package practice.excel.excel;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class _01_ExcelReadTest01 {
	
	public static void main(String[] args) throws Exception{
		
		String filePath = "C:\\Users\\pine\\Desktop\\OT\\4월2주\\1_smart_city\\excel\\[2016_2022생명과학1_화학1]오답_통계.xlsx";
		FileInputStream file = new FileInputStream(filePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getPhysicalNumberOfRows();
		
		int rowNo = 0;
		int cellIndex = 0;
		
		for(rowNo = 0; rowNo < rows; rowNo++) {
			XSSFRow row = sheet.getRow(rowNo);
			
			if(row != null) {
				
				int cells = row.getPhysicalNumberOfCells();
				
				for(cellIndex = 0; cellIndex <= cells ; cellIndex++) {
					
					XSSFCell cell = row.getCell(cellIndex);
					
					if(cell == null) {
						continue;
					} else {
						String value = "";
						
						switch (cell.getCellType()) {
						case FORMULA:
							value = cell.getCellFormula();
							break;
							
						case NUMERIC:
							value = cell.getNumericCellValue() + "";
							break;
							
						case STRING:
							value = cell.getStringCellValue();
							break;
							
						case BOOLEAN:
							value = cell.getBooleanCellValue() + "";
							break;
							
						case BLANK:
							value = cell.getBooleanCellValue() + "";
							break;
							
						case ERROR:
							value = cell.getErrorCellString() + "";
							break;
						}
						
						System.out.println(value);
					}
				}
			}
		}		
	}
}
