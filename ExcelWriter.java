package Excel;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {
	String fileOne;
	String fileTwo;
	String fileThree;

	File file;
	Scanner input = new Scanner(System.in);
	XSSFWorkbook xssfWorkbook;

	public ArrayList<String> fileReader(File file) throws IOException {

		BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
		String read = " ";
		ArrayList<String> listOne = new ArrayList<>();
		while ((read = bufferedReader.readLine()) != null) {
			listOne.add(read);
		}
		return listOne;
	}

	public void writeToXlsxFile() throws IOException, FileNotFoundException {

		String userDirectoryString = System.getProperty("user.dir");

		System.out.println("Enter the first file name");
		fileOne = input.nextLine();
		String firstPathString = userDirectoryString + File.separator + fileOne;
		file = new File(firstPathString);
		ArrayList<String> firstFile = fileReader(file);
		System.out.println("Student one details : " + firstFile);

		System.out.println("Enter the second file name");
		fileTwo = input.nextLine();
		String scondPathString = userDirectoryString + File.separator + fileTwo;
		file = new File(scondPathString);
		ArrayList<String> secondFile = fileReader(file);
		System.out.println("Student two details : " + secondFile);

		System.out.println("Enter the third file");
		fileThree = input.nextLine();
		String thirdPathString = userDirectoryString + File.separator + fileThree;
		file = new File(thirdPathString);
		ArrayList<String> thirdFile = fileReader(file);
		System.out.println("Student three details : " + thirdFile);

		ArrayList<String> header = new ArrayList<>();
		header.add("RollNo");
		header.add("Name");
		header.add("Marks");
		header.add("Status");
		header.add("Degree");

		System.out.println("Enter your FileName");
		String filename = input.next();

		String excelPath = userDirectoryString + File.separator + filename;
		File file = new File(excelPath);

		FileInputStream fileInputStream = new FileInputStream(file);
		xssfWorkbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet3");
		XSSFRow nRow;

		Map<String, ArrayList> map = new TreeMap<>();
		map.put("0", header);
		map.put("1", firstFile);
		map.put("2", secondFile);
		map.put("3", thirdFile);

		Set<String> key = map.keySet();
		int row = 0; // Null pointer Exception
		for (String str : key) {
			nRow = xssfSheet.createRow(row++);
			ArrayList objArray = map.get(str);
			int cell = 0;
			for (Object obj : objArray) {
				XSSFCell xssfCell = nRow.createCell(cell++);
				xssfCell.setCellValue((String) obj);
			}

		}
		FileOutputStream fileOutputStream = new FileOutputStream(file);
		xssfWorkbook.write(fileOutputStream);
		fileOutputStream.close();
		System.out.println("XLsheet was writed Successfully");

	}

	public static void main(String[] args) throws IOException {
		ExcelWriter excelWriter = new ExcelWriter();
		excelWriter.writeToXlsxFile();
	}

}