package Excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	// Scanner input = new Scanner(System.in);
	XSSFWorkbook xssfWorkbook;
	// String filePath = "./ExcelSheet/Task.xlsx";
	File file = new File("C:\\Users\\santhoshkumar.s\\eclipse-workspace\\Automation\\ExcelSheet\\Task.xlsx");

	public void writeExcel() throws IOException {
//		System.out.println("Enter Your File Name:");
//		String fileName = input.nextLine();
		File file = new File("C:\\Users\\santhoshkumar.s\\eclipse-workspace\\Automation\\ExcelSheet\\Task.xlsx");

		FileInputStream fileInputStream = new FileInputStream(file);
		xssfWorkbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet4");

		XSSFRow nRow;

		Object[] headerObject = new Object[] { "Name", "Team" };
		Object[] firstObject = new Object[] { "Sasi", "Falcon" };
		Object[] SecondObject = new Object[] { "Lokesh", "COE" };
		Object[] thirdObject = new Object[] { "Akilan", "CoE" };
		Object[] finalObject = new Object[] { "uDHAY", "WK" };

		LinkedHashMap<String, Object[]> map = new LinkedHashMap<>();
		map.put("0", headerObject);
		map.put("1", firstObject);
		map.put("2", SecondObject);
		map.put("3", thirdObject);
		map.put("4", finalObject);

		Set<String> key = map.keySet();
		int row = 0; // Null pointer Exception

		for (String str : key) {
			nRow = xssfSheet.createRow(row++);
			Object[] objArray = map.get(str);

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

	public void readExcel() throws IOException, InvalidFormatException {
//		System.out.println("Enter Your File Name:");
//		String filePath = input.nextLine();

		xssfWorkbook = new XSSFWorkbook(file);
		XSSFSheet xssfSheet = xssfWorkbook.getSheet("Sheet4");
		// XSSFSheet xssfSheet1 = xssfWorkbook.getSheetAt(2);

		int lastRowNum = xssfSheet.getLastRowNum();
		int physicalNumRows = xssfSheet.getPhysicalNumberOfRows();
		int cellNum = xssfSheet.getRow(1).getLastCellNum();
		int numberOfCells = 0;

		System.out.println("Number of cells in each row :" + cellNum);
		System.out.println("Inclusive of Header Row count : " + physicalNumRows);
		System.out.println("Total Number of rows : " + lastRowNum);

		for (int rowCount = 0; rowCount < lastRowNum; rowCount++) {
			XSSFRow xssfRow = xssfSheet.getRow(rowCount);
			for (int cellCount = 0; cellCount < cellNum; cellCount++) {
				XSSFCell xssfCell = xssfRow.getCell(cellCount);

				// String data = xssfCell.getStringCellValue();
				DataFormatter dataFormatter = new DataFormatter();
				String data = dataFormatter.formatCellValue(xssfCell);

				System.out.print(data + "   ");

				numberOfCells++;
			}
			System.out.println();
		}
		System.out.println("Number of cells availed with values wr to rowCount :" + numberOfCells);

	}

	public static void main(String[] args) throws IOException, InvalidFormatException {
		Excel handlingExcel = new Excel();
		handlingExcel.writeExcel();
		handlingExcel.readExcel();
	}
}
