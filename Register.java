package Projects;

import java.io.File;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.Reporter;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.thoughtworks.selenium.webdriven.WebDriverBackedSelenium;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;

public class Register 
{
public static WebDriver driver;
public static Workbook wb;
public Sheet sh;
static WebDriverBackedSelenium selenium;
static String str;
	@Test
	public static void Registration() throws BiffException, IOException, InterruptedException
	{
		File fi=new File("D:\\manga\\register.xls");
		wb=Workbook.getWorkbook(fi);
		Sheet ws=wb.getSheet(0);
		for(int i=1;i<ws.getRows();i++)
			{
			driver.findElement(By.name("firstName")).sendKeys(ws.getCell(0, i).getContents());
			driver.findElement(By.name("lastName")).sendKeys(ws.getCell(1, i).getContents());
			driver.findElement(By.name("phone")).sendKeys(ws.getCell(2, i).getContents());
			driver.findElement(By.name("userName")).sendKeys(ws.getCell(3, i).getContents());
			driver.findElement(By.name("address1")).sendKeys(ws.getCell(4, i).getContents());
			driver.findElement(By.name("city")).sendKeys(ws.getCell(5, i).getContents());
			driver.findElement(By.name("state")).sendKeys(ws.getCell(6, i).getContents());
			driver.findElement(By.name("postalCode")).sendKeys(ws.getCell(7, i).getContents());
			driver.findElement(By.name("country")).sendKeys(ws.getCell(8, i).getContents());
			driver.findElement(By.name("email")).sendKeys(ws.getCell(9, i).getContents());
			driver.findElement(By.name("password")).sendKeys(ws.getCell(10, i).getContents());
			driver.findElement(By.name("confirmPassword")).sendKeys(ws.getCell(11, i).getContents());
			driver.findElement(By.name("register")).click();
			Thread.sleep(3000);
			if(selenium.isElementPresent("link=SIGN-OFF"))
			{
			driver.findElement(By.linkText("SIGN-OFF")).click();
			Thread.sleep(1000);
			driver.get("http://newtours.demoaut.com");
			System.out.println("user login is success");
			Reporter.log("both user and password are valid");
			str="Pass";
			}else{
				str="Fail";
				System.out.println("user login is unsuccess");
				Reporter.log("both user and password are invalid");
					}
		// Add the str value to Result file under result column	
			Label result=new Label(2, i, str);
			sh.addCell(result);

			}
		//Add 3 labels in Result file
				  Label rs=new Label(2,0,"Result");
				   sh.addCell(rs);
			// Write and close result file	  
				  wwb.write();
				  wwb.close();
			}
		
	}
	@BeforeTest
	public static void login()
	{
		driver=new FirefoxDriver();
		selenium=new WebDriverBackedSelenium(driver, "http://newtours.demoaut.com");
		driver.get("http://newtours.demoaut.com");
		driver.findElement(By.linkText("REGISTER")).click();
	}
	@AfterTest
	public static void Logout()
	{
		driver.quit();
	}
	}
