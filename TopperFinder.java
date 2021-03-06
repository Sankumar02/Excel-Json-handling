package JSON;
//Write a Java program to perform JSON File Reading, JSON File Writing, and Excel File Writing. 

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class TopperFinder {
	String pathString;
	Scanner oScanner = new Scanner(System.in);
	JSONObject writer;
	JSONArray arrayWrite = new JSONArray();
	XSSFWorkbook xssfWorkbook;
	File file;
	ArrayList<String> highestValue = new ArrayList<>();

	public void writeToJsonFile() throws IOException, FileNotFoundException {

		System.out.println("Enter the file name : ");

		String fileName = oScanner.nextLine();
		String userDirectory = System.getProperty("user.dir");
		pathString = userDirectory + File.separator + fileName;

		System.out.println("Enter the count of student : ");
		int count = oScanner.nextInt();
		for (int i = 0; i < count; i++) {
			System.out.println("Enter the Roll number : ");
			String rollNumber = oScanner.next();

			System.out.println("Enter the Name : ");
			String name = oScanner.next();

			System.out.println("Enter the Total Marks : ");
			String totalMarks = oScanner.next();

			System.out.println("Enter the Result : ");
			String result = oScanner.next();

			writer = new JSONObject();
			writer.put("Roll Number", rollNumber);
			writer.put("Name", name);
			writer.put("Total Marks", totalMarks);
			writer.put("Result", result);

			arrayWrite.add(writer);
		}
		FileWriter writeFile = new FileWriter(pathString);
		writeFile.write(arrayWrite.toJSONString());
		writeFile.close();
		System.out.println("JSON file has been created !!");

	}

	public ArrayList<String> readJsonFile(int countObject) throws IOException, ParseException {
		JSONParser jsonParser = new JSONParser();
		JSONArray array = new JSONArray();
		ArrayList<String> arrayList;

		FileReader readFile = new FileReader(pathString);

		Object objParse = jsonParser.parse(readFile);
		array = (JSONArray) objParse;

		JSONObject readJsonObject = (JSONObject) array.get(countObject);

		String rollNumber = (String) readJsonObject.get("Roll Number");

		String name = (String) readJsonObject.get("Name");

		String totalMarks = (String) readJsonObject.get("Total Marks");

		String result = (String) readJsonObject.get("Result");

		arrayList = new ArrayList<>();
		arrayList.add(rollNumber);
		arrayList.add(name);
		arrayList.add(totalMarks);
		arrayList.add(result);
		return arrayList;

	}

	public void findTheTopperStudent() throws IOException, ParseException {
		ArrayList<String> firstArrayList = readJsonFile(0);
		ArrayList<String> secondArrayList = readJsonFile(1);
		ArrayList<String> thirdArrayList = readJsonFile(2);
		ArrayList<String> fourthArrayList = readJsonFile(3);
		ArrayList<String> fifthArrayList = readJsonFile(4);

		System.out.println("Student one Details : " + firstArrayList);
		System.out.println("Student two details : " + secondArrayList);
		System.out.println("Student three details : " + thirdArrayList);
		System.out.println("Student fourth details : " + fourthArrayList);
		System.out.println("Student fifth details : " + fifthArrayList);

		String list1 = firstArrayList.get(2);
		int firstValue = Integer.parseInt(list1);
		String list2 = secondArrayList.get(2);
		int secondValue = Integer.parseInt(list2);
		String list3 = thirdArrayList.get(2);
		int thirdValue = Integer.parseInt(list3);
		String list4 = fourthArrayList.get(2);
		int fourthValue = Integer.parseInt(list4);
		String list5 = fifthArrayList.get(2);
		int fifthValue = Integer.parseInt(list5);

		TreeSet<String> set = new TreeSet<>();
		set.add(list1);
		set.add(list2);
		set.add(list3);
		set.add(list4);
		set.add(list5);
		String highestMark = set.last();
		int value = Integer.parseInt(highestMark);
		System.out.println("Higher mark scored among the three students : " + highestMark);

		if (value == firstValue) {
			highestValue = firstArrayList;
		} else if (value == secondValue) {
			highestValue = secondArrayList;
		} else if (value == thirdValue) {
			highestValue = thirdArrayList;
		} else if (value == fourthValue) {
			highestValue = fourthArrayList;
		} else if (value == fifthValue) {
			highestValue = fifthArrayList;
		}
	}

	public void writetoExcel() throws IOException {

		ArrayList<String> header = new ArrayList<>();
		header.add("Roll Number");
		header.add("Name");
		header.add("Total Marks");
		header.add("Result");

		System.out.println("Enter your FileName");
		String excelFileName = oScanner.next();
		String userDirectoryExcel = System.getProperty("user.dir");
		String excelPath = userDirectoryExcel + File.separator + excelFileName;
		File excelFile = new File(excelPath);

		FileInputStream fileInputStream = new FileInputStream(excelFile);
		xssfWorkbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet6");
		XSSFRow nRow;

		Map<String, ArrayList> map = new TreeMap<>();
		map.put("0", header);
		map.put("1", highestValue);

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
