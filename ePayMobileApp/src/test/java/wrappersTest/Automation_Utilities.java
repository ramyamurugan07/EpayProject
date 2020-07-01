package wrappersTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.ScreenOrientation;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import io.appium.java_client.MobileBy;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;

public class Automation_Utilities implements MainWrappers {

	public AndroidDriver<AndroidElement> driver;
	protected static Properties prop;

	protected String productNameInPDP = null;
	protected String productDescriptionInPDP = null;
	protected String productPriceInPDP = null;

	protected String productNameInCartPage = null;
	protected String productDescriptionInCartPage = null;
	protected String productPriceInCartPage = null;

	public void LanchApp() throws MalformedURLException {
		System.out.println("Inside app");
		DesiredCapabilities cap = new DesiredCapabilities();
		
		String sAppName = "Amazon_shopping.apk"; 
		
		cap.setCapability("deviceName", "ae97ca77"); 
		cap.setCapability("platformName", "Android");
		
		cap.setCapability("app",System.getProperty("user.dir") + "/app/" + sAppName); 
		cap.setCapability("appPackage", "com.amazon.mShop.android.shopping");
		cap.setCapability("appActivity", "com.amazon.mShop.home.HomeActivity");
		cap.setCapability("noReset", true);
		driver = new AndroidDriver<AndroidElement>(new URL("http://0.0.0.0:4723/wd/hub"), cap);
		driver.manage().timeouts().implicitlyWait(30, java.util.concurrent.TimeUnit.SECONDS);

		//Screen Orientation check and Set
		ScreenOrientation sOrientation = driver.getOrientation();
		System.out.println("By Default: "+sOrientation.value());
		if(sOrientation.value().contains("landscape"))
		{
			driver.rotate(ScreenOrientation.PORTRAIT);
			System.out.println("Orientation Changed to: "+driver.getOrientation().value());
		}
	}

	//Click Element By ID
	public void clickByID(String sID) {
		driver.findElementById(sID).click();
		WebDriverWait wait = new WebDriverWait(driver, 30);
		AndroidElement eSearchElement = (AndroidElement) wait.until(ExpectedConditions.elementToBeClickable(MobileBy.id(sID)));
		eSearchElement.click();
	}
	
	//Click Element by Xpath
	public void clickByXPath(String sXpath) {
		WebDriverWait wait = new WebDriverWait(driver, 30);
		AndroidElement eSearchElement = (AndroidElement) wait.until(ExpectedConditions.elementToBeClickable(MobileBy.xpath(sXpath)));
		eSearchElement.click();
	}
	
	//Click Element by AccessibilityId
	public void clickByAccessibilityID(String sID){
		WebDriverWait wait = new WebDriverWait(driver, 30);
		AndroidElement eSearchElement = (AndroidElement) wait.until(ExpectedConditions.elementToBeClickable(MobileBy.AccessibilityId(sID)));
		eSearchElement.click();
	}
	//Send the value by using IB
	public void typeByID(String sID, String sValue) {
		WebDriverWait wait = new WebDriverWait(driver, 30);
		AndroidElement eSearchElement = (AndroidElement) wait.until(ExpectedConditions.elementToBeClickable(MobileBy.id(sID)));
		eSearchElement.clear();
		eSearchElement.sendKeys(sValue);
	}
	//Get Element by ID
	public String verifyAndGetByID(String sID) {
		driver.findElementById(sID).isDisplayed();
		String val = driver.findElementById(sID).getText();
		return val;
	}
	//Get Element by Xpath
	public String verifyAndGetByPath(String sXpath){
		driver.findElementByXPath(sXpath).isDisplayed();
		String var = driver.findElementByXPath(sXpath).getText();
		return var;
	}
	//To Check element is selected
	public void verifyByID(String sID) {
		driver.findElementById(sID).isSelected();

	}
	//Validating value by Assertion 
	public void compareValues(String sActual, String sExpected) {
		try{
			Assert.assertEquals(sActual, sExpected);
			System.out.println("Actual value: "+ sActual +" is same as the Expected Value: "+ sExpected );
		} catch (Exception e){
			System.out.println("Actual value: "+ sActual +" is not same as the Expected value: "+ sExpected);
		}
	}
	//Fetch the object.properties
	public void fetchObjects() {
		prop = new Properties();
		try {
			prop.load(new FileInputStream(new File(System.getProperty("user.dir") +"/src/test/java/resources/object.properties")));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	//Take Scrrenshot
	public void takeScreenshot(){
		TakesScreenshot ts = (TakesScreenshot)driver;
		File screenshotFile = ts.getScreenshotAs(OutputType.FILE);
		
		String tm = new SimpleDateFormat("MMDD_mmsss").format(Calendar.getInstance().getTime());
		String sFilepath =System.getProperty("user.dir") + "/Screenshots/" +"500266_"+tm+".png";
		try {
		FileUtils.copyFile(screenshotFile,new File(sFilepath));
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public void closeBrowser(){
		driver.close();
	}

}
