package finalscript;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Rockwell {
	static WebDriver driver;
	static int row_count = 0;
	String[] level_text = new String[14];
	static String sDirPath = System.getProperty("user.dir");

	public void test() {
		String TRC_PART_NUMBER = "'data.csv'!$T$2";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 4, TRC_PART_NUMBER);
		String Level1ProductId = "INDEX('[data.csv]data'!$C$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 9, Level1ProductId);
		String L1Type = "INDEX('[data.csv]data'!$D$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 10, L1Type);
		String Level2ProductId = "INDEX('[data.csv]data'!$E$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 11, Level2ProductId);
		String L2Type = "INDEX('[data.csv]data'!$F$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 12, L2Type);
		String Level3ProductId = "INDEX('[data.csv]data'!$G$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 13, Level3ProductId);
		String L3Type = "INDEX('[data.csv]data'!$H$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 14, L3Type);
		String Level4ProductId = "INDEX('[data.csv]data'!$I$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 15, Level4ProductId);
		String L4Type = "INDEX('[data.csv]data'!$J$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 16, L4Type);
		String Level5ProductId = "INDEX('[data.csv]data'!$K$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 17, Level5ProductId);
		String L5Type = "INDEX('[data.csv]data'!$L$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 18, L5Type);
		String FriendlyURL = "INDEX('[data.csv]data'!$M$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 19, FriendlyURL);
		String NavigationText = "INDEX('[data.csv]data'!$N$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 20, NavigationText);
		String PageHeading = "INDEX('[data.csv]data'!$O$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 21, PageHeading);
		String SubcatListingimage = "INDEX('[data.csv]data'!$P$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 22, SubcatListingimage);
		String ShortDescriptionImage = "INDEX('[data.csv]data'!$Q$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 23, ShortDescriptionImage);
		String OverviewTabImage = "INDEX('[data.csv]data'!$R$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 24, OverviewTabImage);
		String ProductDescription = "INDEX('[data.csv]data'!$S$2,MATCH(E2,'[data.csv]data'!$T$2,0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 25, ProductDescription);
	}

	@BeforeTest
	public void configuration() {
		FileOutputStream fileOut1 = null;
		FileOutputStream fileOut2 = null;

		// To delete old data.csv and final.xlsx files
		try {
			File file = new File(sDirPath + "\\output\\data.csv");
			File file1 = new File(sDirPath + "\\output\\Final.xlsx");
			file.delete();
			file1.delete();

		} catch (Exception e) {
			e.printStackTrace();
		}
		// To create old data.csv and final.xlsx files
		try {
			fileOut1 = new FileOutputStream(sDirPath + "\\output\\data.csv");
			Workbook wb = new XSSFWorkbook();
			FileOutputStream fileOut = new FileOutputStream(sDirPath + "\\output\\Final.xlsx");
			Sheet sheet1 = wb.createSheet("Sheet1");
			Row row = sheet1.createRow((short) 0);
			for (int i = 0; i < 26; i++) {
				Cell c = row.createCell(i);
			}
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {

			e.printStackTrace();
		}
		// Populating Headers in data.csv and final.xlsx files
		writeToCsv("Status", false);
		writeToCsv("Pri/ Sec", false);
		writeToCsv("Level 1 Product Id", false);
		writeToCsv("L1 Type", false);
		writeToCsv("Level 2 Product Id", false);
		writeToCsv("L2 Type", false);
		writeToCsv("Level 3 Product Id", false);
		writeToCsv("L3 Type", false);
		writeToCsv("Level 4 Product Id", false);
		writeToCsv("L4 Type", false);
		writeToCsv("Level 5 Product Id", false);
		writeToCsv("L5 Type", false);
		writeToCsv("Level 6 Product Id", false);
		// writeToCsv("Level 7 Product Id", false);
		writeToCsv("Friendly URL", false);
		writeToCsv("Navigation Text", false);
		writeToCsv("Page Heading", false);
		writeToCsv("Subcat Listing image", false);
		writeToCsv("Short Description Image", false);
		writeToCsv("Overview Tab Image", false);
		writeToCsv("Product Description", false);
		writeToCsv("Product Catalog Number", true);
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 0, "TRC_PART_NUMBER");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 1, "MANUFACTURER_NAME");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 2, "BRAND_NAME");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 3, "UPC");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 4, "MANUFACTURER_PART_NUMBER");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 5, "UNSPSC");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 6,
				"ACTIVE(Active='Y' InActive='N' Planned Obsolescence='P' Obsolete='O' Item Alert='T' Unlisted='U' Non-price Maintained='K' Withdrawn='W' Pending To Delete='X')");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 7, "DISPLAY_ONLINE");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 8, "ORIG_SORT");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 9, "Level 1 Product Id");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 10, "L1 Type");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 11, "Level 2 Product Id");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 12, "L2 Type");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 13, "Level 3 Product Id");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 14, "L3 Type");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 15, "Level 4 Product Id");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 16, "L4 Type");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 17, "Level 5 Product Id");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 18, "L5 Type");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 19, "Friendly URL");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 20, "Navigation Text");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 21, "Page Heading");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 22, "Subcat Listing image");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 23, "Short Description Image");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 24, "Overview Tab Image");
		writeExcel(sDirPath + "\\output\\Final.xlsx", "Sheet1", 0, 25, "Product Description");
		String sheetName = getExcelData(sDirPath + "\\configuration\\congif.xlsx", "config", 1, 2);
		int rowStartNo = getExcelNumericData(sDirPath + "\\configuration\\congif.xlsx", "config", 2, 2);
		int rowEndNo = getExcelNumericData(sDirPath + "\\configuration\\congif.xlsx", "config", 3, 2);
		String TRC_PART_NUMBER = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$A$" + rowStartNo + ":$A$"
				+ rowEndNo + ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$"
				+ rowEndNo + ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 0, TRC_PART_NUMBER);
		String MANUFACTURER_NAME = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$B$" + rowStartNo + ":$B$"
				+ rowEndNo + ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$"
				+ rowEndNo + ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 1, MANUFACTURER_NAME);
		String BRAND_NAME = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$C$" + rowStartNo + ":$C$" + rowEndNo
				+ ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$" + rowEndNo
				+ ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 2, BRAND_NAME);
		String UPC = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$D$" + rowStartNo + ":$D$" + rowEndNo
				+ ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$" + rowEndNo
				+ ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 3, UPC);
		String UNSPSC = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$F$" + rowStartNo + ":$F$" + rowEndNo
				+ ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$" + rowEndNo
				+ ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 5, UNSPSC);
		String ACTIVE = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$G$" + rowStartNo + ":$G$" + rowEndNo
				+ ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$" + rowEndNo
				+ ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 6, ACTIVE);
		String DISPLAY_ONLINE = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$H$" + rowStartNo + ":$H$"
				+ rowEndNo + ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$"
				+ rowEndNo + ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 7, DISPLAY_ONLINE);
		String ORIG_SORT = "INDEX('[distributorInputFile.xlsx]" + sheetName + "'!$I$" + rowStartNo + ":$I$" + rowEndNo
				+ ",MATCH(E2,'[distributorInputFile.xlsx]" + sheetName + "'!$E$" + rowStartNo + ":$E$" + rowEndNo
				+ ",0))";
		writeExcelData(sDirPath + "\\output\\Final.xlsx", "Sheet1", 1, 8, ORIG_SORT);
	}

	public static void writeExcel(String filePath, String sheetName, int rowNo, int cellNo, String data) {
		try {
			FileInputStream fileInput = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fileInput);
			Sheet st = wb.getSheet(sheetName);
			Row r = st.getRow(rowNo);
			if (r == null) {
				r = st.createRow(rowNo);
			}
			Cell c = r.createCell(cellNo);
			c.setCellType(Cell.CELL_TYPE_STRING);
			c.setCellValue(data);
			// CellStyle cs = wb.createCellStyle();
			// c.setCellValue(data);
			FileOutputStream fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void writeExcelData(String filePath, String sheetName, int rowNo, int cellNo, String data) {
		try {
			FileInputStream fileInput = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fileInput);
			Sheet st = wb.getSheet(sheetName);
			Row r = st.getRow(rowNo);
			if (r == null) {
				r = st.createRow(rowNo);
			}
			Cell c = r.createCell(cellNo);
			c.setCellType(Cell.CELL_TYPE_FORMULA);
			// c.setCellFormula(data);
			// CellStyle cs = wb.createCellStyle();
			// c.setCellValue(data);
			FileOutputStream fileOut = new FileOutputStream(filePath);
			wb.write(fileOut);
			// evalExcelData(filePath, sheetName, rowNo, cellNo);
			// XSSFFormulaEvaluator.evaluateAllFormulaCells((XSSFWorkbook) wb);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static int getExcelNumericData(String filePath, String sheetName, int rowNo, int cellNo) {
		try {
			FileInputStream fileInput = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fileInput);
			Sheet st = wb.getSheet(sheetName);
			Row r = st.getRow(rowNo);
			Cell c = r.getCell(cellNo);
			int data = (int) c.getNumericCellValue();
			return data;
		} catch (Exception e) {
			return 0;
		}
	}

	public static String getExcelData(String filePath, String sheetName, int rowNo, int cellNo) {
		try {
			FileInputStream fileInput = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fileInput);
			Sheet st = wb.getSheet(sheetName);
			Row r = st.getRow(rowNo);
			Cell c = r.getCell(cellNo);
			String data = c.getStringCellValue();
			return data;
		} catch (Exception e) {
			return " ";
		}
	}

	public static void evalExcelData(String filePath, String sheetName, int rowNo, int cellNo) {
		try {
			FileInputStream fileInput = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fileInput);
			Sheet st = wb.getSheet(sheetName);
			Row r = st.getRow(rowNo);
			Cell c = r.getCell(cellNo);
			CreationHelper createHelper = wb.getCreationHelper();

			createHelper.createFormulaEvaluator().evaluateAll();
			// createHelper.createFormulaEvaluator().evaluateInCell(c);
			// String data = c.getStringCellValue();
			// return data;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/*
	 * Description: Configuration to launch Application under test with respect to
	 * browser
	 */
	public void launchApp() {
		if (getExcelData(sDirPath + "\\browserConfig.xlsx", "Sheet1", 1, 0).equalsIgnoreCase("Chrome")) {
			System.setProperty("webdriver.chrome.driver", sDirPath + "\\browserDriver\\chromedriver.exe");
			driver = new ChromeDriver();
		} else if (getExcelData(sDirPath + "\\browserConfig.xlsx", "Sheet1", 1, 0).equalsIgnoreCase("firefox")) {
			System.setProperty("webdriver.gecko.driver", sDirPath + "\\browserDriver\\geckodriver.exe");
			driver = new FirefoxDriver();
		} else {
			System.setProperty("webdriver.ie.driver", sDirPath + "\\browserDriver\\IEDriverServer.exe");
			driver = new InternetExplorerDriver();
		}
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("http://ab.rockwellautomation.com/allenbradley/productdirectory.page?");
		driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		if (driver.findElement(By.className("btn")) != null) {
			driver.findElement(By.className("btn")).click();
			driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
			driver.findElement(By.xpath("//ul/li[3]/div/div[2]/a")).click();
		}
		
		System.out.println("git test1");
		System.out.println("git test2");
		System.out.println("git test3");
		System.out.println("git test4");

		// *[@id="top"]/div[3]/div[2]/div[1]/div/div/div/ul/li[3]/div/div[2]/a
		// driver.findElement(By.xpath("//button[contains(text(),'No
		// Preference')]")).click();
	}

	/*
	 * Description: Explicit wait for particular element
	 */
	static void waitDynamicallyforelement(int timeInSec, String waitConditionLocator) {
		WebDriverWait wait = new WebDriverWait(driver, timeInSec);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(waitConditionLocator)));
	}

	/*
	 * Description: Explicit wait for particular multiple elements
	 */
	static void waitDynamicallyforelements(int timeInSec, String waitConditionLocator) {
		WebDriverWait wait = new WebDriverWait(driver, timeInSec);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(waitConditionLocator)));
	}

	@Test(enabled = true, description = "Navigate to each product, Click on product selection tab, populate catalog numbers and description into 'data.csv' file")
	public void productContent() {
		launchApp();
		String product_content = null;
		String fall_back = driver.getCurrentUrl();
		System.out.println(fall_back);
		// Identifying total number of modules
		List<WebElement> element1 = driver.findElements(By.xpath("//div[@id='productdirectory']//h3/a"));
		try {
			int start = 0;
			int condition = 0;
			if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 1, 0).equalsIgnoreCase("yes")) {
				start = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 1, 1);
				condition = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 1, 2);
			} else if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 1, 0).equalsIgnoreCase("no")) {
				start = 1;
				condition = element1.size();
			}
			//debug --ns1
			for(int j=0; j <element1.size();j++) {
				System.out.println("elemtent "+j);
				System.out.println(element1.get(j).getText());
				System.out.println("elemtent size"+ element1.size());
			}
			
			// To iterate each module
			for (int i = start; i <= condition; i++) {
				String st1 = "(//div[@id='productdirectory']//h3/a)[" + i + "]";				
				try {
					WebElement ele = driver.findElement(By.xpath(st1));
					level_text[0] = ele.getText();
					System.out.println(level_text[0]);
					logGenerator(level_text[0]);
					ele.click();
					level_one();
					Thread.sleep(3000);
					driver.get(fall_back);
				} catch (Exception e) {
					logGenerator("-----------------------------Started-----------------------------------------------");
					logGenerator(driver.getCurrentUrl());
					StackTraceElement[] x = e.getStackTrace();
					for (StackTraceElement stackTrace : x) {
						if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
							logGenerator(level_text[0]);
							logGenerator("In class: " + stackTrace.getClassName());
							logGenerator("at line number: " + stackTrace.getLineNumber());
						}
					}
					logGenerator("-----------------------------Ended-------------------------------------------------");
				}
				level_text[0] = null;
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(level_text[0]);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}
	}

	/*
	 * Description: To handle level - 1
	 */
	public void level_one() {
		String fall_back = null;
		String l0_text = null;
		try {
			Thread.sleep(2000);
			fall_back = driver.getCurrentUrl();
			// To identify No of submodules in level - 1
			List<WebElement> li1 = driver.findElements(By.xpath("//div[@id='catsubcatcontent']"));
			Thread.sleep(1000);
			// To iterate each module in level - 1
			int start = 0;
			int condition = 0;
			if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 3, 0).equalsIgnoreCase("yes")) {
				start = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 3, 1);
				condition = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 3, 2);
			} else if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 3, 0).equalsIgnoreCase("no")) {
				start = 1;
				condition = li1.size();
			}
			for (int k = start; k <= condition; k++) {
				Thread.sleep(5000);
				String str1 = "(//div[@id='catsubcatcontent'])[" + k + "]//a";
				List<WebElement> ele = driver.findElements(By.xpath(str1));
				System.out.println("--ns--ele =  " + ele);		
				// To check if the sub module in level - 1 is having sub modules
				if (ele.size() > 2) {
					String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
					System.out.println("--ns--ele size =  " + ele.size());	
					try {
						WebElement element1 = driver.findElement(By.xpath(str2));
						level_text[1] = element1.getText();
						logGenerator("--" + level_text[1]);
						System.out.println("--" + level_text[1]);
						element1.click();
						level_two();
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[1]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					level_text[1] = null;
				} else {
					String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a"; // title
					try {
						WebElement element1 = driver.findElement(By.xpath(str2));
						level_text[1] = element1.getText();
						System.out.println("--" + level_text[1]);
						logGenerator("--" + level_text[1]);
						String link_text = element1.getText();
						String img_url = "(//div[@id='catsubcatcontent'])[" + k + "]//a/img";
						String sub_img_src = driver.findElement(By.xpath(img_url)).getAttribute("src");
						element1.click();
						before_final(sub_img_src, link_text);
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[1]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					level_text[1] = null;
				}
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(level_text[1]);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}
	}

	/*
	 * Description: To handle sub level-2
	 */
	public void level_two() {
		String fall_back = null;
		String l0_text = null;
		try {
			Thread.sleep(2000);
			fall_back = driver.getCurrentUrl();
			// To identify No of submodules in level - 2
			List<WebElement> li1 = driver.findElements(By.xpath("//div[@id='catsubcatcontent']"));
			Thread.sleep(1000);
			// To iterate each module in level - 2
			int start = 0;
			int condition = 0;
			if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 5, 0).equalsIgnoreCase("yes")) {
				start = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 5, 1);
				condition = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 5, 2);
			} else if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 5, 0).equalsIgnoreCase("no")) {
				start = 1;
				condition = li1.size();
			}
			for (int k = start; k <= condition; k++) {
				Thread.sleep(2000);
				String str1 = "(//div[@id='catsubcatcontent'])[" + k + "]//a";
				List<WebElement> ele = driver.findElements(By.xpath(str1));
				// To check if the sub module in level - 2 is having sub modules
				if (ele.size() > 2) {
					String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
					try {
						WebElement element1 = driver.findElement(By.xpath(str2));
						level_text[2] = element1.getText();
						logGenerator("----" + level_text[2]);
						System.out.println("----" + level_text[2]);
						element1.click();
						level_three();
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[2]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					level_text[2] = null;
				} else {
					String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
					try {
						WebElement element1 = driver.findElement(By.xpath(str2));
						String _eleText = driver.findElement(By.xpath(str2)).getText();
						level_text[2] = element1.getText();
						logGenerator("----" + level_text[2]);
						System.out.println("----" + level_text[2]);
						String link_text = element1.getText();
						String img_url = "(//div[@id='catsubcatcontent'])[" + k + "]//a/img";
						String sub_img_src = driver.findElement(By.xpath(img_url)).getAttribute("src");

						/// Summary
						// "Connected Components Workbench Software" is opening
						/// a new window
						// Checking for "Connected Components Workbench
						/// Software" link and skipping this.
						/// End Summary

						// if (_eleText.equals("Connected Components Workbench Software")) {
						// logGenerator("");
						// System.out.println("---- Skipped - Connected Components Workbench Software
						// ");
						// } else {
						// element1.click();
						// }
						element1.click();
						Thread.sleep(2000);
						before_final(sub_img_src, link_text);
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[2]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					level_text[2] = null;
				}
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(level_text[2]);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}

	}

	/*
	 * Description: To handle sub level-3
	 */
	public void level_three() {
		String fall_back = null;
		String l0_text = null;
		String link_text = "";
		String sub_img_src = "";
		try {
			Thread.sleep(2000);
			fall_back = driver.getCurrentUrl();
			// To identify No of submodules in level - 3
			List<WebElement> li1 = driver.findElements(By.xpath("//div[@id='catsubcatcontent']"));
			Thread.sleep(1000);
			// To check if the sub module in present level - 3
			if (li1.size() > 0) {
				// To iterate each module in level - 3

				int start = 0;
				int condition = 0;
				if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 7, 0).equalsIgnoreCase("yes")) {
					start = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 7, 1);
					condition = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 7, 2);
				} else if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 7, 0).equalsIgnoreCase("no")) {
					start = 1;
					condition = li1.size();
				}

				System.out.println("level 3 "+start);
				System.out.println("level 3 "+condition);
				
				for (int k = start; k <= condition; k++) {
					Thread.sleep(2000);
					String str0 = "(//div[@id='catsubcatcontent'])[" + k + "]//a";
					List<WebElement> ele0 = driver.findElements(By.xpath(str0));
					String str1 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
					List<WebElement> ele = driver.findElements(By.xpath(str1));
					// To check if the sub module in level - 3 is having sub modules
					if (ele0.size() > 2) {
						String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
						try {
							WebElement element1 = driver.findElement(By.xpath(str2));
							level_text[3] = element1.getText();
							logGenerator("----" + level_text[3]);
							System.out.println("----" + level_text[3]);
							element1.click();
							level_four();
							Thread.sleep(3000);
							driver.get(fall_back);
						} catch (Exception e) {
							logGenerator(
									"-----------------------------Started-----------------------------------------------");

							logGenerator(driver.getCurrentUrl());
							StackTraceElement[] x = e.getStackTrace();
							for (StackTraceElement stackTrace : x) {
								if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
									logGenerator(level_text[3]);
									logGenerator("In class: " + stackTrace.getClassName());
									logGenerator("at line number: " + stackTrace.getLineNumber());
								}
							}
							logGenerator(
									"-----------------------------Ended-------------------------------------------------");
						}
						level_text[3] = null;
					} else if (ele.size() > 1) // if products are >1
					{
						String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
						try {
							WebElement element1 = driver.findElement(By.xpath(str2));
							level_text[3] = element1.getText();
							logGenerator("------" + level_text[3]);
							System.out.println("------" + level_text[3]);
							element1.click();
							level_four();
							Thread.sleep(3000);
							driver.get(fall_back);
						} catch (Exception e) {
							logGenerator(
									"-----------------------------Started-----------------------------------------------");

							logGenerator(driver.getCurrentUrl());
							StackTraceElement[] x = e.getStackTrace();
							for (StackTraceElement stackTrace : x) {
								if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
									logGenerator(level_text[3]);
									logGenerator("In class: " + stackTrace.getClassName());
									logGenerator("at line number: " + stackTrace.getLineNumber());
								}
							}
							logGenerator(
									"-----------------------------Ended-------------------------------------------------");

						}
						level_text[3] = null;
					} else {
						String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
						try {
							WebElement element1 = driver.findElement(By.xpath(str2));
							level_text[3] = element1.getText();
							logGenerator("------" + level_text[3]);
							System.out.println("------" + level_text[3]);
							link_text = element1.getText();
							String img_url = "(//div[@id='catsubcatcontent'])[" + k + "]//a/img";
							sub_img_src = driver.findElement(By.xpath(img_url)).getAttribute("src");
							element1.click();
							before_final(sub_img_src, link_text);
							Thread.sleep(3000);
							driver.get(fall_back);
						} catch (Exception e) {
							logGenerator(
									"-----------------------------Started-----------------------------------------------");
							logGenerator(driver.getCurrentUrl());
							StackTraceElement[] x = e.getStackTrace();
							for (StackTraceElement stackTrace : x) {
								if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
									logGenerator(level_text[3]);
									logGenerator("In class: " + stackTrace.getClassName());
									logGenerator("at line number: " + stackTrace.getLineNumber());
								}
							}
							logGenerator(
									"-----------------------------Ended-------------------------------------------------");
						}
						level_text[3] = null;
					}
				}
			} else {
				try {
					driver.findElement(By.cssSelector("#products")).click();
					driver.findElement(By.cssSelector("#products")).click();
					driver.findElement(By.cssSelector("#products")).click();
				} catch (Exception e) {
					logGenerator("-----------------------------Started-----------------------------------------------");
					logGenerator(driver.getCurrentUrl());
					StackTraceElement[] x = e.getStackTrace();
					for (StackTraceElement stackTrace : x) {
						if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
							logGenerator(level_text[3]);
							logGenerator("In class: " + stackTrace.getClassName());
							logGenerator("at line number: " + stackTrace.getLineNumber());
						}
					}
					logGenerator("-----------------------------Ended-------------------------------------------------");
				}
				List<WebElement> li2 = driver.findElements(By.xpath("//div[@id='products']//a"));
				fall_back = driver.getCurrentUrl();
				for (int i = 1; i <= li2.size(); i++) {
					String str10 = "(//div[@id='products']//a)[" + i + "]";
					try {
						WebElement element = driver.findElement(By.xpath(str10));
						level_text[3] = element.getText();
						logGenerator("------" + level_text[3]);
						System.out.println("------" + level_text[3]);

						// Create a loop to print cell values in a row
						for (int j = 0; j <= level_text.length; j++) {

							// Print Excel data in console
							System.out.print(level_text[j] + "|| ");
						}
						System.out.println();

						link_text = element.getText();
						element.click();
						before_final(sub_img_src, link_text);
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[3]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					level_text[3] = null;
				}
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(level_text[3]);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}
	}

	/*
	 * Description: To handle sub level-4
	 */
	public void level_four() {
		try {
			Thread.sleep(4000);
		} catch (InterruptedException e1) {
			e1.printStackTrace();
		}
		String fall_back = null;
		String l0_text = null;
		try {
			fall_back = driver.getCurrentUrl();
			// To identify No of submodules in level - 4
			List<WebElement> li1 = driver.findElements(By.xpath("//div[@id='catsubcatcontent']"));
			// To iterate each module in level - 4
			int start = 0;
			int condition = 0;
			if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 9, 0).equalsIgnoreCase("yes")) {
				start = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 9, 1);
				condition = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 9, 2);
			} else if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 9, 0).equalsIgnoreCase("no")) {
				start = 1;
				condition = li1.size();
			}
			
			System.out.println("level four "+start);
			System.out.println("level four "+condition);
			
			for (int k = start; k <= condition; k++) {
				Thread.sleep(2000);
				String str1 = "(//div[@id='catsubcatcontent'])[" + k + "]//a";
				List<WebElement> ele = driver.findElements(By.xpath(str1));
				// String str2 = "(//div[@id='catsubcatcontent'])["+k+"]//h4/a";

				// To check if the sub module in level - 4 is having sub modules
				if (ele.size() > 2) {
					String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
					try {
						WebElement element1 = driver.findElement(By.xpath(str2));
						level_text[4] = element1.getText();
						logGenerator("--" + level_text[4]);
						System.out.println("--" + level_text[4]);
						element1.click();
						Thread.sleep(2000);
						level_five();
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[4]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");

					}
					level_text[4] = null;
				} else {
					try {
						String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
						WebElement element1 = driver.findElement(By.xpath(str2));
						level_text[4] = element1.getText();
						logGenerator("--------" + level_text[4]);
						System.out.println("--------" + level_text[4]);
						String link_text = element1.getText();
						String img_url = "(//div[@id='catsubcatcontent'])[" + k + "]//a/img";
						String sub_img_src = driver.findElement(By.xpath(img_url)).getAttribute("src");
						element1.click();
						before_final(sub_img_src, link_text);
						Thread.sleep(3000);
						driver.get(fall_back);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(level_text[4]);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					level_text[4] = null;
				}
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(level_text[4]);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}
	}

	/*
	 * Description: To handle sub level-5
	 */
	public void level_five() {
		try {
			Thread.sleep(2000);
		} catch (InterruptedException e1) {
			e1.printStackTrace();
		}
		String fall_back = null;
		try {
			fall_back = driver.getCurrentUrl();
			// To identify No of submodules in level - 5
			List<WebElement> li1 = driver.findElements(By.xpath("//div[@id='catsubcatcontent']"));
			Thread.sleep(1000);
			// To iterate each module in level - 5

			int start = 0;
			int condition = 0;
			if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 11, 0).equalsIgnoreCase("yes")) {
				start = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 11, 1);
				condition = getExcelNumericData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 11, 2);
			} else if (getExcelData(sDirPath + "//moduleConfig.xlsx", "Sheet1", 11, 0).equalsIgnoreCase("no")) {
				start = 1;
				condition = li1.size();
			}

			for (int k = start; k <= condition; k++) {
				Thread.sleep(2000);
				try {
					String str2 = "(//div[@id='catsubcatcontent'])[" + k + "]//h4/a";
					WebElement element1 = driver.findElement(By.xpath(str2));
					level_text[5] = element1.getText();
					logGenerator("--------" + level_text[5]);
					System.out.println("--------" + level_text[5]);
					String link_text = element1.getText();
					String img_url = "(//div[@id='catsubcatcontent'])[" + k + "]//a/img";
					String sub_img_src = driver.findElement(By.xpath(img_url)).getAttribute("src");
					element1.click();
					before_final(sub_img_src, link_text);
					Thread.sleep(3000);
					driver.get(fall_back);
				} catch (Exception e) {
					logGenerator("-----------------------------Started-----------------------------------------------");
					logGenerator(driver.getCurrentUrl());
					StackTraceElement[] x = e.getStackTrace();
					for (StackTraceElement stackTrace : x) {
						if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
							logGenerator(level_text[5]);
							logGenerator("In class: " + stackTrace.getClassName());
							logGenerator("at line number: " + stackTrace.getLineNumber());
						}
					}
					logGenerator("-----------------------------Ended-------------------------------------------------");
				}
				level_text[5] = null;
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(level_text[5]);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}
	}

	/*
	 * Description: After clicking on the product, working with product selection
	 * tab
	 */
	public void before_final(String sub_img_src, String link_text) {
		try {
			Thread.sleep(3000);
		} catch (InterruptedException e1) {
			e1.printStackTrace();
		}
		String page_disp_name = "";
		String image1 = "";
		String str_url = "";

		String winHandleBefore = driver.getWindowHandle();
		if (driver.getWindowHandles().size() > 1) {
			Set<String> windowhandles = driver.getWindowHandles();
			for (String windowHandle : windowhandles) {
				driver.switchTo().window(windowHandle);
			}
			logGenerator("******* product is opened in new tab ********");
			logGenerator(driver.getCurrentUrl());
			writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src, "Opened in new tab");
			driver.close();
			driver.switchTo().window(winHandleBefore);
		}
		try {
			page_disp_name = driver.findElement(By.xpath("//div[@id='content-intro']//h1")).getText();
			image1 = driver.findElement(By.cssSelector("#content-intro img.img-responsive")).getAttribute("src");
			str_url = driver.getCurrentUrl();
			try {
				// If product selection tab is present
				if (driver.findElements(By.xpath("//li[@id='selection']")).size() > 0) // tab
				{
					try {
						driver.findElement(By.xpath("//li[@id='selection']")).click();
						Thread.sleep(5000);
					} catch (Exception e) {
						logGenerator(
								"-----------------------------Started-----------------------------------------------");
						logGenerator(driver.getCurrentUrl());
						StackTraceElement[] x = e.getStackTrace();
						for (StackTraceElement stackTrace : x) {
							if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
								logGenerator(link_text);
								logGenerator("In class: " + stackTrace.getClassName());
								logGenerator("at line number: " + stackTrace.getLineNumber());
							}
						}
						logGenerator(
								"-----------------------------Ended-------------------------------------------------");
					}
					try {
						WebElement element = driver.findElement(By.xpath("//div[@id='selection-iframe']/iframe"));
						WebDriverWait wait = new WebDriverWait(driver, 30);
						wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(element));
						/*
						 * if product selection down message is displayed refresh the page
						 * 
						 */
						if (driver.findElements(By.xpath("//h2[contains(text(),'Product Selection site is down')]"))
								.size() > 0) {
							capture(driver, " site1");
							driver.get(driver.getCurrentUrl());
							Thread.sleep(15000);
							driver.switchTo().frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
							if (driver.findElements(By.xpath("//h2[contains(text(),'Product Selection site is down')]"))
									.size() > 0) {
								capture(driver, " site2");
								driver.get(driver.getCurrentUrl());
								Thread.sleep(15000);
								driver.switchTo().frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
								if (driver
										.findElements(
												By.xpath("//h2[contains(text(),'Product Selection site is down')]"))
										.size() > 0) {
									capture(driver, " site3");
									driver.get(driver.getCurrentUrl());
									Thread.sleep(15000);
									driver.switchTo()
									.frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
									if (driver
											.findElements(
													By.xpath("//h2[contains(text(),'Product Selection site is down')]"))
											.size() > 0) {
										capture(driver, " site4");
										driver.get(driver.getCurrentUrl());
										Thread.sleep(15000);
										driver.switchTo()
										.frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
										if (driver
												.findElements(By.xpath(
														"//h2[contains(text(),'Product Selection site is down')]"))
												.size() > 0) {
											capture(driver, " site5");
											driver.get(driver.getCurrentUrl());
											Thread.sleep(15000);
											capture(driver, " site6");
											writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
													"Site is down");
										}
									}
								}
							}
							try {
								// After refreshing if Sub-families are displayed
								if (driver.findElement(By.id("theHeader")).isDisplayed()) {
									// System.out.println("in site down check");
									List<WebElement> li1 = driver.findElements(By.xpath("//ul[@id='simpleUl']/li")); // To
									// identify
									// all
									// the
									// products
									// after
									// clicking
									// on
									// the
									// product
									// selection
									// tab
									for (int l = 1; l <= li1.size(); l++) {
										driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
										int a = 8;
										int count = 2;
										if (l == a) {
											String str = "(//ul[@class='mSPages']/li)[" + count + "]"; // To identify
											// pagination
											// radio buttons
											level_text[6] = driver.findElement(By.xpath(str)).getText();
											driver.findElement(By.xpath(str)).click();
											Thread.sleep(5000);
											// To handle loading image
											if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
												Thread.sleep(2000);
												// System.out.println("In loading image Msg 1");
												if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
													// System.out.println("In loading image Msg 2");
													Thread.sleep(60000);
												}
											}
											if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a"))
													.size() > 0) {
												// waitDynamicallyforelement(240, "(//div[@id='ProductsGrid']//a)[1]");
												if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
														.size() > 0)
													final_level(page_disp_name, image1, str_url, link_text,
															sub_img_src);
											} else if (driver.findElements(By.xpath(
													"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"))
													.size() > 0) {
												// To handle child's under each sub families
												List<WebElement> child = driver.findElements(By.xpath(
														"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"));
												for (int n = 0; n <= child.size(); n++) {
													level_text[6] = child.get(n).getText();
													child.get(n).click();
													Thread.sleep(5000);
													// To handle loading image
													if (driver.findElement(By.id("LoadingImage"))
															.isDisplayed() == true) {
														Thread.sleep(2000);
														// System.out.println("In loading image Msg 1");
														if (driver.findElement(By.id("LoadingImage"))
																.isDisplayed() == true) {
															// System.out.println("In loading image Msg 2");
															Thread.sleep(60000);
														}
													}
													if (driver.findElements(By.xpath("//span[text()='Catalog Number']"))
															.size() > 0) {
														// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
														if (driver
																.findElements(
																		By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
																.size() > 0) {
															final_level(page_disp_name, image1, str_url, link_text,
																	sub_img_src);
														} else {
															writeMessage(page_disp_name, image1, str_url, link_text,
																	sub_img_src, "Data is not loaded");
														}
													}
												}
												driver.findElement(By.xpath(str)).click();
												Thread.sleep(2000);
											}
											count++;
											a = +7;
										}
										String str5 = "(//ul[@id='simpleUl']/li)[" + l + "]";
										level_text[6] = driver.findElement(By.xpath(str5)).getText();
										driver.findElement(By.xpath(str5)).click();
										Thread.sleep(5000);
										// To handle loading image
										if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
											Thread.sleep(2000);
											// System.out.println("In loading image Msg 1");
											if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
												// System.out.println("In loading image Msg 2");
												Thread.sleep(60000);
											}
										}
										if (level_text[6].equalsIgnoreCase("800F Legend Plates")) {
											// System.out.println(" In 800F Legend Plates");
											Thread.sleep(150000);
										}
										if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a")).size() > 0) {
											// waitDynamicallyforelement(240, "(//div[@id='ProductsGrid']//a)[1]");
											if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
													.size() > 0) {
												final_level(page_disp_name, image1, str_url, link_text, sub_img_src);
											} else {
												writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
														"Data is not loaded");
											}
										} else {
											writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
													"No Data to Populate");
										}
										if (driver.findElements(By.xpath(
												"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"))
												.size() > 0) {
											List<WebElement> child = driver.findElements(By.xpath(
													"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"));
											for (int n = 0; n <= child.size() - 1; n++) {
												level_text[6] = child.get(n).getText();
												child.get(n).click();
												Thread.sleep(5000);
												// To handle loading image
												if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
													Thread.sleep(2000);
													// System.out.println("In loading image Msg 1");
													if (driver.findElement(By.id("LoadingImage"))
															.isDisplayed() == true) {
														// System.out.println("In loading image Msg 2");
														Thread.sleep(60000);
													}
												}
												if (driver.findElements(By.xpath("//span[text()='Catalog Number']"))
														.size() > 0) {
													// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
													if (driver
															.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
															.size() > 0) {
														final_level(page_disp_name, image1, str_url, link_text,
																sub_img_src);
													} else {
														writeMessage(page_disp_name, image1, str_url, link_text,
																sub_img_src, "Data is not loaded");
													}
												} else {
													writeMessage(page_disp_name, image1, str_url, link_text,
															sub_img_src, "No Data to Populate");
												}
												level_text[6] = null;
											}
											driver.findElement(By.xpath(str5)).click();
											Thread.sleep(5000);
										}
										level_text[6] = null;
									}
									driver.switchTo().defaultContent();
								} else {
									Thread.sleep(5000);
									if (driver.findElements(By.xpath("//span[text()='Catalog Number']")).size() > 0) {
										// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
										if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
												.size() > 0) {
											final_level(page_disp_name, image1, str_url, link_text, sub_img_src);
										} else {
											writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
													"Data is not loaded");
										}
									} else {
										writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
												"No Data to Populate");
									}
									driver.switchTo().defaultContent();
								}
							} catch (Exception e) {
								logGenerator(
										"-----------------------------Started-----------------------------------------------");
								logGenerator(driver.getCurrentUrl());
								StackTraceElement[] x = e.getStackTrace();
								for (StackTraceElement stackTrace : x) {
									if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
										logGenerator(link_text);
										logGenerator("In class: " + stackTrace.getClassName());
										logGenerator("at line number: " + stackTrace.getLineNumber());
									}
								}
								logGenerator(
										"-----------------------------Ended-------------------------------------------------");
							}
						}
						/*
						 * if No data available message is displayed, refresh the page maximum five
						 * times
						 */
						else if (driver.findElement(By.id("NoData")).isDisplayed()) {
							int h = 0;
							for (int i = 1; i <= 5; i++) {
								driver.get(driver.getCurrentUrl());
								Thread.sleep(8000);
								driver.switchTo().frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
								if (driver
										.findElements(
												By.xpath("//h2[contains(text(),'Product Selection site is down')]"))
										.size() > 0) {
									driver.get(driver.getCurrentUrl());
									Thread.sleep(10000);
									driver.switchTo()
									.frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
									if (driver
											.findElements(
													By.xpath("//h2[contains(text(),'Product Selection site is down')]"))
											.size() > 0) {
										driver.get(driver.getCurrentUrl());
										Thread.sleep(10000);
										driver.switchTo()
										.frame(driver.findElement(By.xpath("//div[@id='selection']//iframe")));
									}
								}
								if (driver.findElements(By.id("NoData")).size() < 0) {
									break;
								}
								h = i;
							}
							if (h == 5) {
								writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
										"No Data Available Message");
							} else {
								try {
									// if Sub-families are displayed
									if (driver.findElement(By.id("theHeader")).isDisplayed()) {
										List<WebElement> li1 = driver.findElements(By.xpath("//ul[@id='simpleUl']/li")); // To
										// identify
										// all
										// the
										// products
										// after
										// clicking
										// on
										// product
										// selection
										// tab
										for (int l = 1; l <= li1.size(); l++) {
											int a = 8;
											int count = 2;
											if (l == a) {
												String str = "(//ul[@class='mSPages']/li)[" + count + "]"; // To
												// identify
												// Pagination
												// radio
												// buttons
												level_text[6] = driver.findElement(By.xpath(str)).getText();
												driver.findElement(By.xpath(str)).click();
												Thread.sleep(5000);
												// To handle loading image
												if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
													Thread.sleep(2000);
													// System.out.println("In loading image Msg 1");
													if (driver.findElement(By.id("LoadingImage"))
															.isDisplayed() == true) {
														// System.out.println("In loading image Msg 2");
														Thread.sleep(60000);
													}
												}
												if (driver.findElements(By.xpath("//span[text()='Catalog Number']"))
														.size() > 0) {
													// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
													if (driver
															.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
															.size() > 0) {
														final_level(page_disp_name, image1, str_url, link_text,
																sub_img_src);
													} else {
														writeMessage(page_disp_name, image1, str_url, link_text,
																sub_img_src, "Data is not loaded");
													}
												} else {
													writeMessage(page_disp_name, image1, str_url, link_text,
															sub_img_src, "No data to populate");
												}
												if (driver.findElements(By.xpath(
														"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"))
														.size() > 0) {// To handle child's under each sub families
													List<WebElement> child = driver.findElements(By.xpath(
															"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"));
													for (int n = 0; n <= child.size(); n++) {
														level_text[6] = child.get(n).getText();
														child.get(n).click();
														Thread.sleep(5000);
														// To handle loading image
														if (driver.findElement(By.id("LoadingImage"))
																.isDisplayed() == true) {
															Thread.sleep(2000);
															// System.out.println("In loading image Msg 1");
															if (driver.findElement(By.id("LoadingImage"))
																	.isDisplayed() == true) {
																// System.out.println("In loading image Msg 2");
																Thread.sleep(60000);
															}
														}
														if (driver
																.findElements(
																		By.xpath("//span[text()='Catalog Number']"))
																.size() > 0) {
															// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
															if (driver
																	.findElements(By
																			.xpath("(//div[@id='ProductsGrid']//a)[1]"))
																	.size() > 0) {
																final_level(page_disp_name, image1, str_url, link_text,
																		sub_img_src);
															} else {
																writeMessage(page_disp_name, image1, str_url, link_text,
																		sub_img_src, "Data is not loaded");
															}
														} else {
															writeMessage(page_disp_name, image1, str_url, link_text,
																	sub_img_src, "No Data to Populate");
														}
													}
													driver.findElement(By.xpath(str)).click();
													Thread.sleep(2000);
												}
												count++;
												a = +7;
											}
											String str5 = "(//ul[@id='simpleUl']/li)[" + l + "]";
											level_text[6] = driver.findElement(By.xpath(str5)).getText();
											driver.findElement(By.xpath(str5)).click();
											Thread.sleep(5000);
											// To handle loading image
											if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
												Thread.sleep(2000);
												// System.out.println("In loading image Msg 1");
												if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
													System.out.println("In loading image Msg 2");
													Thread.sleep(60000);
												}
											}
											if (level_text[6].equalsIgnoreCase("800F Legend Plates")) {
												// System.out.println("800F Legend Plates");
												Thread.sleep(150000);
											}
											if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a"))
													.size() > 0) {
												// waitDynamicallyforelement(240, "(//div[@id='ProductsGrid']//a)[1]");
												if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
														.size() > 0) {
													final_level(page_disp_name, image1, str_url, link_text,
															sub_img_src);
												} else {
													writeMessage(page_disp_name, image1, str_url, link_text,
															sub_img_src, "Data is not loaded");
												}
											} else {
												writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
														"No data to populate");
											}
											if (driver.findElements(By.xpath(
													"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"))
													.size() > 0) {// To handle child's under each sub families
												List<WebElement> child = driver.findElements(By.xpath(
														"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"));
												for (int n = 0; n <= child.size() - 1; n++) {
													level_text[6] = child.get(n).getText();
													child.get(n).click();
													Thread.sleep(5000);
													// To handle loading image
													if (driver.findElement(By.id("LoadingImage"))
															.isDisplayed() == true) {
														Thread.sleep(2000);
														// System.out.println("In loading image Msg 1");
														if (driver.findElement(By.id("LoadingImage"))
																.isDisplayed() == true) {
															// System.out.println("In loading image Msg 2");
															Thread.sleep(60000);
														}
													}
													if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a"))
															.size() > 0) {
														// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
														if (driver
																.findElements(
																		By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
																.size() > 0) {
															final_level(page_disp_name, image1, str_url, link_text,
																	sub_img_src);
														} else {
															writeMessage(page_disp_name, image1, str_url, link_text,
																	sub_img_src, "Data is not loaded");
														}
													} else {
														writeMessage(page_disp_name, image1, str_url, link_text,
																sub_img_src, "No data to populate");
													}
													level_text[6] = null;
												}
												driver.findElement(By.xpath(str5)).click();
												Thread.sleep(2000);
											}
											level_text[6] = null;
										}
										driver.switchTo().defaultContent();
									} else {
										Thread.sleep(5000);
										if (driver.findElements(By.xpath("//span[text()='Catalog Number']"))
												.size() > 0) {
											// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
											if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
													.size() > 0) {
												final_level(page_disp_name, image1, str_url, link_text, sub_img_src);
											} else {
												writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
														"data is not loaded");
											}
										} else {
											writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
													"No Data to Populate");
										}
										driver.switchTo().defaultContent();
									}
								} catch (Exception e) {
									logGenerator(
											"-----------------------------Started-----------------------------------------------");
									logGenerator(driver.getCurrentUrl());
									StackTraceElement[] x = e.getStackTrace();
									for (StackTraceElement stackTrace : x) {
										if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
											logGenerator(link_text);
											logGenerator("In class: " + stackTrace.getClassName());
											logGenerator("at line number: " + stackTrace.getLineNumber());
										}
									}
									logGenerator(
											"-----------------------------Ended-------------------------------------------------");
								}
							}
						} else { // if Sub-families are displayed
							try {
								if (driver.findElement(By.id("theHeader")).isDisplayed()) {
									/*
									 * To identify all the products after clicking on the product selection tab
									 */
									List<WebElement> li1 = driver.findElements(By.xpath("//ul[@id='simpleUl']/li"));
									// System.out.println(li1.size());
									for (int l = 1; l <= li1.size(); l++) {
										// System.out.println("*****"+l+"******");
										driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
										int a = 8;
										int count = 2;
										if (l == a) {
											// System.out.println("in l==a");
											/*
											 * To identify pagination radio buttons
											 */
											String str = "(//ul[@class='mSPages']/li)[" + count + "]";
											level_text[6] = driver.findElement(By.xpath(str)).getText();
											// System.out.println("in l==a "+level_text[6]);
											driver.findElement(By.xpath(str)).click();
											Thread.sleep(5000);
											// To handle loading image
											if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
												Thread.sleep(2000);
												// System.out.println("In loading image Msg 1");
												if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
													// System.out.println("In loading image Msg 2");
													Thread.sleep(60000);
												}
											}
											if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a"))
													.size() > 0) {
												// waitDynamicallyforelement(240, "(//div[@id='ProductsGrid']//a)[1]");
												if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
														.size() > 0) {
													final_level(page_disp_name, image1, str_url, link_text,
															sub_img_src);
												} else {
													writeMessage(page_disp_name, image1, str_url, link_text,
															sub_img_src, "Data is not loaded");
												}
											}
											// To handle child's under each sub families
											else if (driver.findElements(By.xpath(
													"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"))
													.size() > 0) {
												List<WebElement> child = driver.findElements(By.xpath(
														"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"));
												// System.out.println("in l==a chile size "+child.size());
												for (int n = 0; n <= child.size(); n++) {
													level_text[6] = child.get(n).getText();
													// System.out.println(level_text[6]);
													child.get(n).click();
													Thread.sleep(5000);
													// To handle loading image
													if (driver.findElement(By.id("LoadingImage"))
															.isDisplayed() == true) {
														Thread.sleep(2000);
														// System.out.println("In loading image Msg 1");
														if (driver.findElement(By.id("LoadingImage"))
																.isDisplayed() == true) {
															// System.out.println("In loading image Msg 2");
															Thread.sleep(60000);
														}
													}
													if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a"))
															.size() > 0) {
														// waitDynamicallyforelement(240,"(//div[@id='ProductsGrid']//a)[1]");
														if (driver
																.findElements(
																		By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
																.size() > 0) {
															final_level(page_disp_name, image1, str_url, link_text,
																	sub_img_src);
														} else {
															writeMessage(page_disp_name, image1, str_url, link_text,
																	sub_img_src, "Data is not loaded");
														}
													} else {
														writeMessage(page_disp_name, image1, str_url, link_text,
																sub_img_src, "No data to populate");
													}
												}
												driver.findElement(By.xpath(str)).click();
												// System.out.println("in l==a
												// reclick--"+driver.findElement(By.xpath(str)).getText());
												Thread.sleep(2000);
											}
											count++;
											a = +7;
										}
										/*
										 * To identify each product in the available products after clicking on the
										 * product selection tab
										 */
										String str5 = "(//ul[@id='simpleUl']/li)[" + l + "]";
										level_text[6] = driver.findElement(By.xpath(str5)).getText();
										// System.out.println(level_text[6]);
										driver.findElement(By.xpath(str5)).click();
										// JavascriptExecutor js = (JavascriptExecutor) driver;
										// js.executeScript("arguments[0].click();",
										// driver.findElement(By.xpath(str5)));
										Thread.sleep(5000);
										// To handle loading image
										if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
											Thread.sleep(2000);
											// System.out.println("In loading image Msg 1");
											if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
												// System.out.println("In loading image Msg 2");
												Thread.sleep(60000);
											}
										}
										if (level_text[6].equalsIgnoreCase("800F Legend Plates")) {
											// System.out.println(" in 800F Legend Plates");
											Thread.sleep(150000);
										}
										// System.out.println("check for grid size
										// "+driver.findElements(By.xpath("//div[@id='ProductsGrid']//a")).size() );
										// System.out.println("check for child size
										// "+driver.findElements(By.xpath("//ul[@id='simpleUl']/li[contains(@class,'selectionChild
										// mSSlide')]")).size());
										if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a")).size() > 0) {
											waitDynamicallyforelement(60, "(//div[@id='ProductsGrid']//a)[1]");
											if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
													.size() > 0) {
												final_level(page_disp_name, image1, str_url, link_text, sub_img_src);
											} else {
												writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
														"Data is not loaded");
											}
										} else {
											writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
													"No data to populate");
										}
										if (driver.findElements(By.xpath(
												"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"))
												.size() > 0) {
											List<WebElement> child = driver.findElements(By.xpath(
													"//ul[@id='simpleUl']/li[contains(@class,'selectionChild mSSlide')]"));
											// System.out.println("child size"+child.size());
											for (int n = 0; n <= child.size() - 1; n++) {
												level_text[6] = child.get(n).getText();
												// System.out.println("Name: "+level_text[6]);
												child.get(n).click();
												Thread.sleep(5000);
												// To handle loading image
												if (driver.findElement(By.id("LoadingImage")).isDisplayed() == true) {
													Thread.sleep(2000);
													// System.out.println("In loading image Msg 1");
													if (driver.findElement(By.id("LoadingImage"))
															.isDisplayed() == true) {
														// System.out.println("In loading image Msg 2");
														Thread.sleep(60000);
													}
												}
												if (driver.findElements(By.xpath("//span[text()='Catalog Number']"))
														.size() > 0) {
													// waitDynamicallyforelement(240,
													// "(//div[@id='ProductsGrid']//a)[1]");
													if (driver
															.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
															.size() > 0) {
														final_level(page_disp_name, image1, str_url, link_text,
																sub_img_src);
													} else {
														writeMessage(page_disp_name, image1, str_url, link_text,
																sub_img_src, "Data is not loaded");
													}
												}
												level_text[6] = null;
											}
											driver.findElement(By.xpath(str5)).click();
											// System.out.println("reclick--"+driver.findElement(By.xpath(str5)).getText());
											Thread.sleep(2000);
										}
										level_text[6] = null;
										Thread.sleep(2000);
									}
									driver.switchTo().defaultContent();
								} else {
									if (driver.findElements(By.xpath("//div[@id='ProductsGrid']//a")).size() > 0) {
										// waitDynamicallyforelement(240, "(//div[@id='ProductsGrid']//a)[1]");
										if (driver.findElements(By.xpath("(//div[@id='ProductsGrid']//a)[1]"))
												.size() > 0) {
											final_level(page_disp_name, image1, str_url, link_text, sub_img_src);
										} else {
											writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
													"Data is not loaded");
										}
									} else {
										writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
												"No data to populate");
									}
									driver.switchTo().defaultContent();
								}
							} catch (Exception e) {
								logGenerator(
										"-----------------------------Started-----------------------------------------------");
								logGenerator(driver.getCurrentUrl());
								StackTraceElement[] x = e.getStackTrace();
								for (StackTraceElement stackTrace : x) {
									if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
										logGenerator(link_text);
										logGenerator("In class: " + stackTrace.getClassName());
										logGenerator("at line number: " + stackTrace.getLineNumber());
									}
								}
								logGenerator(
										"-----------------------------Ended-------------------------------------------------");
							}
						}
					} catch (Exception e) {
						if (driver.findElements(By.xpath("//a[@title='Use the Product Configuration Assistant']"))
								.size() > 0) {
							logGenerator("Product Configuration Assistant is displayed. Product catalog not displayed");
							writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src,
									"Production Configuration");
						}
						driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
					}
				} else {
					writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src, "No Product Selection Tab");
				}
			} catch (Exception e) {
				logGenerator("-----------------------------Started-----------------------------------------------");
				logGenerator(driver.getCurrentUrl());
				StackTraceElement[] x = e.getStackTrace();
				for (StackTraceElement stackTrace : x) {
					if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
						logGenerator(link_text);
						logGenerator("In class: " + stackTrace.getClassName());
						logGenerator("at line number: " + stackTrace.getLineNumber());
					}
				}
				logGenerator("-----------------------------Ended-------------------------------------------------");
			}
		} catch (Exception e) {
			logGenerator("-----------------------------Started-----------------------------------------------");
			logGenerator(driver.getCurrentUrl());
			StackTraceElement[] x = e.getStackTrace();
			for (StackTraceElement stackTrace : x) {
				if (stackTrace.getClassName().equals("FinalScript.Rockwell")) {
					logGenerator(link_text);
					logGenerator("In class: " + stackTrace.getClassName());
					logGenerator("at line number: " + stackTrace.getLineNumber());
				}
			}
			logGenerator("-----------------------------Ended-------------------------------------------------");
		}
	}

	/*
	 * Description: To write level details, description and catalogs into csv file
	 */
	void final_level(String page_disp_name, String image1, String str_url, String link_text, String sub_img_src) {
		try {
			int i = 0;
			WebDriverWait wait1 = new WebDriverWait(driver, 240);
			wait1.until(ExpectedConditions.presenceOfElementLocated(
					By.xpath("//div[@id='ProductsGrid']//a[@title='Display Product Summary']")));
			List<WebElement> product_id = driver
					.findElements(By.xpath("//div[@id='ProductsGrid']//a[@title='Display Product Summary']"));
			List<WebElement> description = driver
					.findElements(By.xpath("(//div[@id='ProductsGrid']//td[@role='gridcell'][2])"));
			// System.out.println("Catalog size " + product_id.size());
			// System.out.println("desc size " + description.size());
			if (product_id.size() == 0) {
				writeMessage(page_disp_name, image1, str_url, link_text, sub_img_src, "No Data to Populate");
			} else {
				do {
					writeToCsv("PROD", false);
					writeToCsv("P", false);
					writeToCsv(level_text[0].replaceAll(",", "@@@@"), false);
					writeToCsv("C", false);
					writeToCsv(level_text[1].replaceAll(",", "@@@@"), false);
					if (level_text[4] == null) {
						if (level_text[3] == null) {
							if (level_text[2] == null) {
								writeToCsv("T", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
							} else {
								writeToCsv("SC", false);
								writeToCsv(level_text[2].replaceAll(",", "@@@@"), false);
								writeToCsv("T", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
								writeToCsv(" ", false);
							}
						}
					}
					if (level_text[4] == null) {
						if (level_text[3] == null) {
						} else {
							writeToCsv("SC", false);
							writeToCsv(level_text[2].replaceAll(",", "@@@@"), false);
							writeToCsv("SSC", false);
							writeToCsv(level_text[3].replaceAll(",", "@@@@"), false);
							writeToCsv("T", false);
							writeToCsv(" ", false);
							writeToCsv(" ", false);
						}
					}
					if (level_text[4] == null) {
					} else {
						writeToCsv("SC", false);
						writeToCsv(level_text[2].replaceAll(",", "@@@@"), false);
						writeToCsv("SSC", false);
						writeToCsv(level_text[3].replaceAll(",", "@@@@"), false);
						writeToCsv("SSSC", false);
						writeToCsv(level_text[4].replaceAll(",", "@@@@"), false);
						writeToCsv("T", false);
					}
					if (level_text[5] == null) {
						writeToCsv(" ", false);
					} else {
						writeToCsv(level_text[5].replaceAll(",", "@@@@"), false);
					}

					// if(level_text[6]==null){
					// writeToCsv(" ", false);
					// }
					// else
					// {
					// writeToCsv(level_text[6].replaceAll(",", "@@@@"), false);
					// }
					writeToCsv(str_url.replaceFirst("http://ab.rockwellautomation.com", ""), false);
					writeToCsv(link_text.replaceAll(",", "@@@@"), false);
					writeToCsv(page_disp_name.replaceAll(",", "@@@@"), false);
					writeToCsv(sub_img_src.replaceFirst("http://ab.rockwellautomation.com", ""), false);
					writeToCsv(image1.replaceFirst("http://ab.rockwellautomation.com", ""), false);
					writeToCsv("", false);
					// description
					String str3 = description.get(i).getText();// description
					// System.out.println("Description " + str3);
					if (str3.length() < 1) {
						Thread.sleep(30000);
						String str3New = description.get(i).getText();
						// System.out.println("Description new "+str3New);
						if (str3New.length() < 1) {

						} else {
							writeToCsv(str3New.replaceAll(",", "@@@@"), false);
						}
					} else {
						writeToCsv(str3.replaceAll(",", "@@@@"), false);
					}

					String str1 = product_id.get(i).getText();// catalogs
					// System.out.println("Catalogs " + str1);
					if (str1.length() < 1) {
						Thread.sleep(30000);
						String str1New = product_id.get(i).getText();
						// System.out.println("catalog New "+str1New);
						if (str1New.length() < 1) {
						} else {
							writeToCsv(str1New.replaceAll(",", "@@@@"), true);
						}
					} else {
						writeToCsv(str1.replaceAll(",", "@@@@"), true);
					}

					i++;
				} while (i < product_id.size());
				row_count = i;
			}
		} catch (Exception e) {
			// writeToCsv("", true);
		}
	}

	/*
	 * Description: To write into csv file
	 */
	public void writeToCsv(String val, boolean finalVal) {
		BufferedWriter b = null;
		try {
			b = new BufferedWriter(new FileWriter(new File(".\\output\\data.csv"), true));
			b.write(val + ",");
			if (finalVal) {
				b.write("\r\n");
			}
			b.close();
			if (b != null) {
				b.close();
			}
		} catch (IOException e) {
			logGenerator(e.getMessage());
		}
	}

	/*
	 * Description: To write into txt file
	 */
	public void logGenerator(String data) {
		try {
			FileWriter fileWrite = new FileWriter("./log.txt", true);
			BufferedWriter bw = new BufferedWriter(fileWrite);
			bw.write(data);
			bw.newLine();
			bw.close();
		} catch (IOException e) {
		}
	}

	/*
	 * Description: To write message under catalog and description columns
	 */
	void writeMessage(String page_dis_name, String img1, String str_Url, String link_txt, String sub_img_Src,
			String message) {
		writeToCsv("PROD", false);
		writeToCsv("P", false);
		writeToCsv(level_text[0].replaceAll(",", "@@@@"), false);
		writeToCsv("C", false);
		writeToCsv(level_text[1].replaceAll(",", "@@@@"), false);
		if (level_text[4] == null) {
			if (level_text[3] == null) {
				if (level_text[2] == null) {
					writeToCsv("T", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
				} else {
					writeToCsv("SC", false);
					writeToCsv(level_text[2].replaceAll(",", "@@@@"), false);
					writeToCsv("T", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
					writeToCsv(" ", false);
				}
			}
		}
		if (level_text[4] == null) {
			if (level_text[3] == null) {
			} else {
				writeToCsv("SC", false);
				writeToCsv(level_text[2].replaceAll(",", "@@@@"), false);
				writeToCsv("SSC", false);
				writeToCsv(level_text[3].replaceAll(",", "@@@@"), false);
				writeToCsv("T", false);
				writeToCsv(" ", false);
				writeToCsv(" ", false);
			}
		}
		if (level_text[4] == null) {
		} else {
			writeToCsv("SC", false);
			writeToCsv(level_text[2].replaceAll(",", "@@@@"), false);
			writeToCsv("SSC", false);
			writeToCsv(level_text[3].replaceAll(",", "@@@@"), false);
			writeToCsv("SSSC", false);
			writeToCsv(level_text[4].replaceAll(",", "@@@@"), false);
			writeToCsv("T", false);
		}
		if (level_text[5] == null) {
			writeToCsv(" ", false);
		} else {
			writeToCsv(level_text[5].replaceAll(",", "@@@@"), false);
		}
		// if(level_text[6]==null){
		// writeToCsv(" ", false);
		// }
		// else
		// {
		// writeToCsv(level_text[6].replaceAll(",", "@@@@"), false);
		// }
		writeToCsv(str_Url.replaceFirst("http://ab.rockwellautomation.com", ""), false);
		writeToCsv(link_txt.replaceAll(",", "@@@@"), false);
		writeToCsv(page_dis_name.replaceAll(",", "@@@@"), false);
		writeToCsv(sub_img_Src.replaceFirst("http://ab.rockwellautomation.com", ""), false);
		writeToCsv(img1.replaceFirst("http://ab.rockwellautomation.com", ""), false);
		writeToCsv("", false);
		writeToCsv(message, false);
		writeToCsv(message, true);
	}

	public static void capture(WebDriver driver, String name) throws IOException {
		Date date = new Date();
		SimpleDateFormat sdf = new SimpleDateFormat("dd-hh_mm_ss");
		String sdate = sdf.format(date);
		TakesScreenshot ts = (TakesScreenshot) driver;
		File source = ts.getScreenshotAs(OutputType.FILE);
		String dest = sDirPath + "\\screenshots\\" + sdate + name + ".png";
		File destination = new File(dest);
		FileUtils.copyFile(source, destination);
	}

	@AfterClass
	public void tearDown() {
		driver.quit();
	}

}