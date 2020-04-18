package testcases;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class SeleniumCode {

	static WebDriver driver;

	public static void launchApp() {

		WebDriverManager.chromedriver().setup();

		driver = new ChromeDriver();

		driver.get("file:///D:/GREENSTECH/Desktop/table.html");

		driver.manage().window().maximize();

	}

	public static List<List<String>> getCompleteTable() {

		launchApp();

		List<WebElement> elements;

		List<WebElement> rows = driver.findElements(By.xpath("//*[@name='BookTable']//tr"));

		List<List<String>> table = new ArrayList<List<String>>();

		List<String> row = new ArrayList<String>();

		for (int i = 0; i < rows.size(); i++) {

			if (i < 1) {

				elements = rows.get(i).findElements(By.tagName("th"));

			}

			else {

				elements = rows.get(i).findElements(By.tagName("td"));

			}

			for (int j = 0; j < elements.size(); j++) {

				row.add(elements.get(j).getText());

			}

			table.add(row);

			row = new ArrayList<String>();

		}

		return table;

	}

	public static void writeExcel() {
		
		List<List<String>> completeTable = getCompleteTable();

		Workbook wb = new XSSFWorkbook();

		Sheet sheet = wb.createSheet("WebTable");

		for (int i = 0; i < completeTable.size(); i++) {
			
			Row row = sheet.createRow(i);
			
			List<String> list = completeTable.get(i);
			
			for (int j = 0; j < completeTable.get(i).size(); j++) {
				
				row.createCell(j).setCellValue(list.get(j));
				
			}

		}

		try (OutputStream os = new FileOutputStream("./WorkBookRepo/Demo.xlsx")) {

			wb.write(os);

		}

		catch (Exception e) {

			System.out.println("File not");

		}

	}

	public static void main(String[] args) {
		
		writeExcel();


	}

}
