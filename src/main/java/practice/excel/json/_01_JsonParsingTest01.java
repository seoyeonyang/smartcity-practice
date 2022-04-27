package practice.excel.json;

import java.io.IOException;

import org.json.JSONObject;

import practice.excel.util.JsonParsingUtil;

public class _01_JsonParsingTest01 {
	
	public static void main(String[] args) throws IOException {
		
		String filePath = "C:/Users/pine/Desktop/OT/4월2주/1_smart_city/json/sample01.json";
		
		JSONObject json = JsonParsingUtil.jsonFileToJsonObject(filePath);
		
		System.out.println(json);
	}

}
