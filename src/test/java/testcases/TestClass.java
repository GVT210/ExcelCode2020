package testcases;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestClass {

	public static void main(String[] args) throws Throwable, IOException {

		Workbook wb = null;

		try (InputStream fileInput = new FileInputStream("D:\\ExcelSession\\sample.xlsx")) {

			wb = new XSSFWorkbook(fileInput);

		}

		catch (Exception e) {

			System.out.println("File not found exception");

		}

		Sheet sheet = wb.getSheetAt(0);

		Row row = sheet.getRow(1);

		Cell cell = row.getCell(1);

		if (cell.getStringCellValue().equals("Amod")) {

			cell.setCellValue("Arjun");
		}

		System.out.println("Done");
		
		wb.write(new FileOutputStream("D:\\ExcelSession\\sample.xlsx"));
		/* DataFormatter format = new DataFormatter(); */

		/*
		 * for (int i = 0; i < noOfRows; i++) {
		 * 
		 * Row row = sheet.getRow(i);
		 * 
		 * int noOfCells = row.getPhysicalNumberOfCells();
		 * 
		 * for (int j = 0; j < noOfCells; j++) {
		 * 
		 * Cell cell = row.getCell(j);
		 * 
		 * switch (cell.getCellType()) {
		 * 
		 * case STRING:
		 * 
		 * System.out.print(cell.getStringCellValue() + "\t");
		 * 
		 * break;
		 * 
		 * case NUMERIC:
		 * 
		 * if (DateUtil.isCellDateFormatted(cell)) {
		 * 
		 * System.out.print(cell.getDateCellValue() + "\t"); } else {
		 * 
		 * System.out.print(cell.getNumericCellValue() + "\t"); }
		 * 
		 * break;
		 * 
		 * case BOOLEAN:
		 * 
		 * System.out.print(cell.getBooleanCellValue());
		 * 
		 * break;
		 * 
		 * case FORMULA:
		 * 
		 * System.out.print(cell.getCellFormula());
		 * 
		 * break;
		 * 
		 * case BLANK:
		 * 
		 * System.out.print(" ");
		 * 
		 * break;
		 * 
		 * default:
		 * 
		 * System.out.print(" None of them matching");
		 * 
		 * }
		 * 
		 * 
		 * String str = format.formatCellValue(row.getCell(j));
		 * 
		 * System.out.println(str);
		 * 
		 * 
		 * }
		 * 
		 * System.out.println();
		 * 
		 * }
		 */
		/*
		 * for (Row row : sheet)
		 * 
		 * {
		 * 
		 * for (Cell cell : row) {
		 * 
		 * System.out.print(cell + "\t");
		 * 
		 * }
		 * 
		 * System.out.println(); }
		 */

		/*
		 * Workbook wb = new XSSFWorkbook();
		 * 
		 * CreationHelper createHelper = wb.getCreationHelper();
		 * 
		 * short format = createHelper.createDataFormat().getFormat("m/d/yy h:mm");
		 * 
		 * Sheet sheet = wb.createSheet("New Sheet");
		 * 
		 * Row row = sheet.createRow(0);
		 * 
		 * row.createCell(0).setCellValue(100);
		 * 
		 * row.createCell(1).setCellValue(true);
		 * 
		 * row.createCell(2).setCellValue(15.45);
		 * 
		 * row.createCell(3).setCellValue("I am a string");
		 * 
		 * CellStyle cellstyle = wb.createCellStyle();
		 * 
		 * cellstyle.setDataFormat(format);
		 * 
		 * row.createCell(4).setCellStyle(cellstyle);
		 * 
		 * row.getCell(4).setCellValue(new Date());
		 * 
		 * row.createCell(5).setCellStyle(cellstyle);
		 * 
		 * row.getCell(5).setCellValue(Calendar.getInstance());
		 * 
		 * try (OutputStream fileOutput = new
		 * FileOutputStream("./WorkBookRepo/Sample.xlsx")) {
		 * 
		 * wb.write(fileOutput);
		 * 
		 * }
		 * 
		 * catch (Exception e) {
		 * 
		 * System.out.println("File not found exception"); }
		 * 
		 * System.out.println("Done");
		 */

	}
}
