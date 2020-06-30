package com.airhacks;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.appium.java_client.MobileBy;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import io.appium.java_client.remote.MobileCapabilityType;

public class Automation_Utilities {
	static String tableViewClassName = "XCUIElementTypeTable";
	Excel_Utilities Obj_Excel_Utilities = new Excel_Utilities();
	Main Obj_Main = new Main();
	//Email Obj_Email = new Email();
	public void Click_By_XPath(AndroidDriver<AndroidElement> driver, String sXpath) {
		Obj_Main.trace("Click_By_XPath() Started - " + sXpath, "LOG");
		WebDriverWait wait = new WebDriverWait(driver, 30);
		WebElement Control = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(sXpath)));
		
		Control.click();		
		Obj_Main.trace("Click_By_XPath() Finished - " + sXpath, "LOG");

	}
	public void Click_By_ID(AndroidDriver<AndroidElement> driver, String sID) throws InterruptedException {
		Obj_Main.trace("Click_By_ID() Started - " + sID, "LOG");
		System.out.println(sID);
		WebDriverWait wait = new WebDriverWait(driver, 30);
		WebElement Control = wait.until(ExpectedConditions.elementToBeClickable(By.id(sID)));
		Control.click();
		Obj_Main.trace("Click_By_ID() Started - " + sID, "LOG");
	}

	public void Click_By_Text(AndroidDriver<AndroidElement> driver, String sText) {
		Obj_Main.trace("Click_By_Text() Started - " + sText, "LOG");
		WebDriverWait wait = new WebDriverWait(driver, 30);
		//WebElement Control = wait.until(ExpectedConditions.elementToBeClickable(By.name(sText)));
		driver.findElement(By.name(sText)).click();
		//Control.click();
		Obj_Main.trace("Click_By_Text() Finished - " + sText, "LOG");
	}

	public void Type_By_xpath(AndroidDriver<AndroidElement> driver, String sXpath, String sText) {

		Obj_Main.trace("Type_By_xpath() Started - " + sXpath + " , " + sText, "LOG");
		WebDriverWait wait = new WebDriverWait(driver, 30);
		WebElement Control = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(sXpath)));
		
		Control.click();
		Control.clear();
		Control.sendKeys(sText);
		Obj_Main.trace("Type_By_xpath() Finished - " + sXpath + " , " + sText, "LOG");

	}

	public void Type_By_ID(AndroidDriver<AndroidElement> driver, String sID, String sText) {
		Obj_Main.trace("Type_By_ID() Started - " + sID + " , " + sText, "LOG");
		WebDriverWait wait = new WebDriverWait(driver, 30);
		WebElement Control = wait.until(ExpectedConditions.elementToBeClickable(By.id(sID)));
		
		Control.click();
		Control.sendKeys(sText);
		Obj_Main.trace("Type_By_ID() Finished - " + sID + " , " + sText, "LOG");

	}
	public void Click_By_ClassName(AndroidDriver<AndroidElement> driver, String sClassName) {
		Obj_Main.trace("Click_By_ClassName() Started - " + sClassName, "LOG");
		driver.findElement(By.className(sClassName)).click();
		Obj_Main.trace("Click_By_ClassName() Finished - " + sClassName, "LOG");
	}
	public void Type_By_ClassName(AndroidDriver<AndroidElement> driver, String sClassName,String sValue) {
		Obj_Main.trace("Click_By_ClassName() Started - " + sClassName, "LOG");
		driver.findElement(By.className(sClassName)).sendKeys(sValue);
		Obj_Main.trace("Click_By_ClassName() Finished - " + sClassName, "LOG");
	}
	public String get_text(AndroidDriver<AndroidElement> driver, String sTextClassName,String sInputClassname) {
		Obj_Main.trace("Click_By_ClassName() Started - " + sTextClassName, "LOG");
		String sValue = driver.findElement(By.className(sTextClassName)).getText();
		Type_By_ClassName(driver, sInputClassname, sValue);
		Obj_Main.trace("Click_By_ClassName() Finished - " + sTextClassName, "LOG");
		return sValue;
		
	}
	public void Click_by_Index(AndroidDriver<AndroidElement> driver, String sTextClassName,String sIndex) {
		Obj_Main.trace("Click_By_ClassName() Started - " + sTextClassName, "LOG");
		int iIndex = Integer.parseInt(sIndex);
		List<AndroidElement> columns1 = driver.findElements(By.className(sTextClassName));
		columns1.get(iIndex).click();
		Obj_Main.trace("Click_By_ClassName() Finished - " + sTextClassName, "LOG");
		
		
	}
	public boolean validation(AndroidDriver<AndroidElement> driver, String sTextClassName,String sExpectedvalue) throws InterruptedException {
		String sValue = driver.findElement(By.className(sTextClassName)).getText();
		
		if (sValue.equalsIgnoreCase(sExpectedvalue)) {
			Assert.assertEquals(sValue, sExpectedvalue);
			return true;
		} else {
			Assert.assertNotEquals(sExpectedvalue, sValue);
			return false;
		}

	}
	public void Scroll_To(AndroidDriver<AndroidElement> driver , String sClassName)
			throws InterruptedException {

		Obj_Main.trace("Scroll_To() Started " + sClassName , "LOG");
		try {

			Thread.sleep(4000);
			RemoteWebElement parent = driver.findElement(By.className(sClassName));
			String parentID = parent.getId();
			Map<String, Object> scrollObject = new HashMap<String, Object>();
			scrollObject.put("element", parentID);
			driver.executeScript("mobile: scroll", scrollObject);
		} catch (Exception e) {
			Obj_Main.trace(e.getLocalizedMessage(), "BOTH");
		}

		Obj_Main.trace("Scroll_To() Finished " + sClassName , "LOG");
	}

	public AndroidDriver<AndroidElement> LaunchApp(String sAppName) throws InterruptedException, IOException{

		Obj_Main.trace("LaunchApp() Started - "+sAppName, "BOTH");
		FileInputStream File = null;
		XSSFWorkbook WorkBook = null;
		XSSFSheet  Config_Sheet = null;
		File = Obj_Excel_Utilities.GetExcel("/Input.xlsm");
		WorkBook = Obj_Excel_Utilities.GetWorkbook(File);
		Config_Sheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, "Config");
		DesiredCapabilities cap = new DesiredCapabilities();
		
		cap.setCapability("appiumVersion", "12.1");
		cap.setCapability(MobileCapabilityType.DEVICE_NAME, Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 20, 14));
		cap.setCapability("deviceOrientation", "portrait");
		//cap.setCapability("browserName", "Android");
		cap.setCapability("platformVersion", Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 20, 16));
		cap.setCapability(MobileCapabilityType.PLATFORM_NAME, Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 20, 15));
		//cap.setCapability("app","com.amazon.mShop.android.shopping");
		sAppName = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 23, 16);
		cap.setCapability("app",System.getProperty("user.dir") + "/app/" + sAppName);
		cap.setCapability("noReset", true);
		cap.setCapability("appPackage", "com.amazon.mShop.android.shopping");
		cap.setCapability("appActivity", "com.amazon.mShop.home.HomeActivity");
		AndroidDriver<AndroidElement> Driver = new AndroidDriver<AndroidElement>(new URL("http://127.0.0.1:4723/wd/hub"), cap);
		//AndroidDriver Driver = new AndroidDriver(new URL("http://127.0.0.1:4723/wd/hub"), cap);
		Thread.sleep(2000);
		
		
		Obj_Main.trace("App launched Successfully", "BOTH");
		Obj_Main.trace("LaunchApp()  Finished - "+sAppName, "BOTH");

		return Driver;
	}
	public void TakeScreenshot(AndroidDriver<AndroidElement> driver, String sFolderPath, String sScreenshotFileName)
			throws IOException, AWTException, InterruptedException

	{
		Obj_Main.trace("TakeScreenshot() Started - " + sFolderPath + " , " + sScreenshotFileName, "LOG");
		String filename = sFolderPath + "/" + sScreenshotFileName;

		Thread.sleep(3000);
		if (driver != null) {

			File screenshotFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			String screenshotBase64 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.BASE64);
			org.apache.commons.io.FileUtils.copyFile(screenshotFile,
					new File(System.getProperty("user.dir") + filename));

		} else {
			Robot robotClassObject = new Robot();
			Rectangle screenSize = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
			BufferedImage tmp = robotClassObject.createScreenCapture(screenSize);
			String path = System.getProperty("user.dir") + "/ScreenCapturesPNG/" + "Auto_Drive_"
					+ System.currentTimeMillis() + ".jpg";
			ImageIO.write(tmp, "jpg", new File(System.getProperty("user.dir") + filename));

		}

		Obj_Main.trace("TakeScreenshot() Finished - " + sFolderPath + " , " + sScreenshotFileName, "LOG");

	}
}
