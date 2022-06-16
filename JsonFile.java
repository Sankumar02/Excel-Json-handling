package JSON;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class JsonFile {
	String patString;
	Scanner oScanner = new Scanner(System.in);
	File file;
	XSSFWorkbook xssfWorkbook;
	String experience;
	String companyName;

	@SuppressWarnings("unchecked")
	public void jsonwriter() throws FileNotFoundException {
		JSONObject writer = new JSONObject();

		writer.put("Employee ID", "2044");
		writer.put("Employee Name", "Santhosh");
		writer.put("Employee Role", "Software Engineer - Trainee");

		JSONArray writeArray = new JSONArray();
		writeArray.add("06 months");
		JSONArray writeCompanyName = new JSONArray();
		writeCompanyName.add("Atmecs");

		writer.put("Experience", writeArray);
		writer.put("Company Name", writeCompanyName);

		try {

			System.out.println("Enter the file name : ");
			String fileName = oScanner.nextLine();
			String userDirectory = System.getProperty("user.dir");
			patString = userDirectory + File.separator + fileName;
			FileWriter fileWrite = new FileWriter(patString);
			fileWrite.write(writer.toJSONString());
			fileWrite.close();
			System.out.println("File writed successfully");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public ArrayList<String> fileReader(File file) throws IOException {

		BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
		String read = " ";
		ArrayList<String> listOne = new ArrayList<>();
		while ((read = bufferedReader.readLine()) != null) {
			listOne.add(read);
		}
		return listOne;
	}

	public void jsonReader() throws ParseException, IOException {
		JSONParser jsonParser = new JSONParser();
		JSONObject oJsonObject = new JSONObject();

		try {
			FileReader readFile = new FileReader(patString);
			Object objParse = jsonParser.parse(readFile);
			oJsonObject = (JSONObject) objParse;
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		String employeeId = (String) oJsonObject.get("Employee ID");
		System.out.println("Employee id : " + employeeId);
		String employeeName = (String) oJsonObject.get("Employee Name");
		System.out.println("Employee name : " + employeeName);
		String employeeRole = (String) oJsonObject.get("Employee Role");
		System.out.println("Employee role : " + employeeRole);

		JSONArray readArray = (JSONArray) oJsonObject.get("Experience");
		Iterator iterator = readArray.iterator();
		while (iterator.hasNext()) {
			experience = (String) iterator.next();
			System.out.println("Experience : " + experience);
		}

		JSONArray readCompanyName = (JSONArray) oJsonObject.get("Company Name");
		Iterator iterateCompanyName = readCompanyName.iterator();
		while (iterateCompanyName.hasNext()) {
			companyName = (String) iterateCompanyName.next();
			System.out.println("Company Name : " + companyName);
		}

		file = new File(patString);
		ArrayList<String> arrayFile = fileReader(file);

		ArrayList<String> header = new ArrayList<>();
		header.add("Employee ID");
		header.add("Employee Name");
		header.add("Employee Role");
		header.add("Experience");
		header.add("Company Name");

		ArrayList<String> details = new ArrayList<String>();
		details.add(employeeId);
		details.add(employeeName);
		details.add(employeeRole);
		details.add(experience);
		details.add(companyName);

		System.out.println("Enter your FileName");
		String excelFileName = oScanner.next();
		String userDirectoryExcel = System.getProperty("user.dir");
		String excelPath = userDirectoryExcel + File.separator + excelFileName;
		File excelFile = new File(excelPath);

		FileInputStream fileInputStream = new FileInputStream(excelFile);
		xssfWorkbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet7");
		XSSFRow nRow;

		Map<String, ArrayList> map = new TreeMap<>();
		map.put("0", header);
		map.put("1", details);

		Set<String> key = map.keySet();
		int row = 0;
		for (String str : key) {
			nRow = xssfSheet.createRow(row++);
			ArrayList objArray = map.get(str);
			int cell = 0;
			for (Object obj : objArray) {
				XSSFCell xssfCell = nRow.createCell(cell++);
				xssfCell.setCellValue((String) obj);
			}

		}
		FileOutputStream fileOutputStream = new FileOutputStream(excelFile);
		xssfWorkbook.write(fileOutputStream);
		fileOutputStream.close();
		System.out.println("XLsheet was writed Successfully");

	}
}
