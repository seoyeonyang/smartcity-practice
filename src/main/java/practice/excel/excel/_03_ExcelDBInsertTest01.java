package practice.excel.excel;

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

public class _03_ExcelDBInsertTest01 {

	public static void main(String[] args) throws Exception {

		//String[] headers = { "examYear", "month", "organization", "wrongCount" };
		String[] headers = { "examYear", "month", "organization", "number", "wrongCount" };

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

					if (cell == null) { 
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

						jobj.put(headers[cellIndex], value);
					}

				} 

				parsedData.add(jobj);

			} 

		}

		String driverName = "org.postgresql.Driver";
		String connectionUrl = "jdbc:postgresql://192.168.1.162:5432/postgres";
		String user = "postgres";
		String password = "0000";

		Connection conn = null;
		PreparedStatement psmt = null;

		try {
			Class.forName(driverName);
			conn = DriverManager.getConnection(connectionUrl, user, password);
			
			String sql = "INSERT INTO sc.exam_statistics(exam_year, exam_month, organization, problem_number, wrong_count) "
					+ " VALUES(?, ?, ?, ?, ?) ";
	
			psmt = conn.prepareStatement(sql);

			int insertedCount = 0;

			for(int i = 0 ; i < parsedData.size() ; i++) {
			
			JSONObject tmpObj = parsedData.get(i);
			
			psmt.setString(1, tmpObj.getString("examYear"));
			psmt.setString(2, tmpObj.getString("month"));
			psmt.setString(3, tmpObj.getString("organization"));
			psmt.setString(4, tmpObj.getString("number"));
			psmt.setString(5, tmpObj.getString("wrongCount"));
	
			insertedCount += psmt.executeUpdate();
		}
		
		System.out.println("insertedCount : " + insertedCount);
		
			
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		}finally {
			if(psmt != null) {psmt.close();}
			if(conn != null) {conn.close();}
		}

	}

}