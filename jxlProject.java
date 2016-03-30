package Projects;

import java.io.File;

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
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class jxlProject 
{
	//public class login {
			public WebDriver driver;
			public WebDriverBackedSelenium selenium;
			public String str;
		 
		@SuppressWarnings("deprecation")
		@Test
		  public void Weblogin() throws Exception{
		//Reading the data from testdata file	  
		File fi=new File("D:\\manga\\login.xls");
				  Workbook w=Workbook.getWorkbook(fi);
				  Sheet s=w.getSheet(0);
		// Creating Result file in Result folder
			  WritableWorkbook wwb=Workbook.createWorkbook(fi,w);
					WritableSheet sh=wwb.getSheet(0);
		for (int i = 1; i < s.getRows(); i++) {
		//Enter username, Password and click on signin by taking data from testdata file	
		driver.findElement(By.name("userName")).sendKeys(s.getCell(0, i).getContents());
		driver.findElement(By.name("password")).sendKeys(s.getCell(1, i).getContents());
		driver.findElement(By.name("login")).click();
			Thread.sleep(1000);
		//Validate signout, if available assign Pass to str, else assign Fail to str	
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
		  @BeforeTest
		  public void beforeTest() {
		
				driver=new FirefoxDriver();
		selenium=new WebDriverBackedSelenium(driver, "http://newtours.demoaut.com");
				driver.get("http://newtours.demoaut.com");
		  }
		  @AfterTest
		  public void afterTest() 
		  {
			 driver.quit();
		  }
	}


