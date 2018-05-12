/*Parameterization in Selenium webDriver using Apache POI(Poor Obfuscation implementation)*/


package testngpackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class ReadWriteExcel {
  WebDriver driver;
  WebDriverWait wait;
  HSSFWorkbook workbook;
  HSSFSheet sheet;
  HSSFCell cell;
  
  @Test
  public void ReadData() throws IOException, WebDriverException {
	  File src = new File("C:\\Users\\Rasika\\Desktop\\Test_Data.xls");
	  //Load the file
	  FileInputStream finput = new FileInputStream(src);
	  //Load the workbook
	  workbook = new HSSFWorkbook(finput);
	  //Load the sheet 
	  sheet = workbook.getSheetAt(0);
	  String Rows= Integer.toString(sheet.getLastRowNum());
	  System.out.println("Last Row number in sheet"+ Rows);
	  for(int i=1; i<=sheet.getLastRowNum();i++)
	  {
		  System.out.println("Now i is "+ i);
		  if (i == 6)
		  {
			  break;
		  }
			  
		  //Import data for First Name
		  cell = sheet.getRow(i).getCell(1);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("Fname")).sendKeys(cell.getStringCellValue());
		//Import data for Last Name
		  cell = sheet.getRow(i).getCell(2);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("Lname")).sendKeys(cell.getStringCellValue());
		//Import data for username
		  cell = sheet.getRow(i).getCell(3);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("username")).sendKeys(cell.getStringCellValue());
		//Import data for Age
		  cell = sheet.getRow(i).getCell(4);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("age")).sendKeys(cell.getStringCellValue());
		//Import data for Email Id
		  cell = sheet.getRow(i).getCell(5);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("email")).sendKeys(cell.getStringCellValue());
		//Import data for Password
		  cell = sheet.getRow(i).getCell(6);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("password")).sendKeys(cell.getStringCellValue());
		//Import data for Re-enter Password
		  cell = sheet.getRow(i).getCell(7);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("password1")).sendKeys(cell.getStringCellValue());
		  //CLick on Sign Up button
		  driver.findElement(By.xpath("/html/body/div/form/div/div/div[9]/button")).click();
		  //Click on Ok Button
		  driver.findElement(By.xpath("/html/body/div[2]/div/div[4]/div/button")).click();
		  
		  wait = new WebDriverWait(driver,30);
		  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		  
		  
		  driver.findElement(By.xpath("/html/body/div/form/div/a")).click();
		  
		  wait = new WebDriverWait(driver,30);
		  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		  
		//Import data for username for login page
		  cell = sheet.getRow(i).getCell(3);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("username")).sendKeys(cell.getStringCellValue());
		  //Import data for password for login page
		  cell=sheet.getRow(i).getCell(6);
		  cell.setCellType(Cell.CELL_TYPE_STRING);
		  driver.findElement(By.name("password")).sendKeys(cell.getStringCellValue());
		  //Click on Login Button
		  driver.findElement(By.xpath("/html/body/div/form/div[1]/div/div[4]/button")).click();
		  //Click on Logout button
		  driver.findElement(By.xpath("/html/body/ul/li[2]/a")).click();
		  //Click on Register Patient Button
		  driver.findElement(By.xpath("/html/body/ul/li[2]/a")).click();
		  //Write data into excel
		  //FileOutputStream foutput = new FileOutputStream(src);
		  //Specify the message that should be present in the cell
		  String message = "Data imported successfully";
		  //Create a cell where data needs to be written
		  sheet.getRow(i).createCell(8).setCellValue(message);
		  // Specify the file in which data needs to be written
		  FileOutputStream fileOutput = new FileOutputStream(src);
		  //finally write content
		  workbook.write(fileOutput);
		  //close file
		  fileOutput.close();
	}
	  driver.close();
	  
  }
  
  @BeforeTest
  public void TestSetup() {
	  System.setProperty("webdriver.chrome.driver","C://Users//Rasika//Desktop//chromedriver.exe");
	  driver = new ChromeDriver();
	  driver.get("http://localhost/Refactoring/EmploymentProject/Register_Employee.php");
	  driver.manage().window().maximize();
	  wait = new WebDriverWait(driver,30);
	  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
  }
  
}
