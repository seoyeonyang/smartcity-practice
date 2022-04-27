package practice.excel;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

public class _02_ExcelParsingTest01 {

	public static void main(String[] args) throws Exception {

		String[] headers = { "examYear", "month", "organization", "wrongCount" };

		String filePath = "C:\\Users\\pine\\Desktop\\OT\\4월2주\\1_smart_city\\excel\\[2016_2022생명과학1_화학1]오답_통계.xlsx";
		FileInputStream file = new FileInputStream(filePath);

		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int rows = sheet.getPhysicalNumberOfRows();

		int rowNo = 0;
		int cellIndex = 0;

		List<JSONObject> parsedData = new ArrayList<JSONObject>();

		for (rowNo = 0; rowNo < rows; rowNo++) {

			XSSFRow row = sheet.getRow(rowNo);

			JSONObject jobj = new JSONObject();

			if (rowNo == 0) {
				continue;
			}

			if (rowNo != 0) {

				int cells = row.getPhysicalNumberOfCells();

				for (cellIndex = 0; cellIndex <= cells; cellIndex++) {

					XSSFCell cell = row.getCell(cellIndex);

					if (cell == null) { // 빈 셀 체크
						continue;
					} else {

						// 타입 별로 내용을 읽는다.
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

						jobj.put(headers[cellIndex], value);
					}

				} // end of for cellIndex

				System.out.println(jobj);
				parsedData.add(jobj);

			} // end of for rowIndex

			System.out.println(parsedData);
		}

	}

}
