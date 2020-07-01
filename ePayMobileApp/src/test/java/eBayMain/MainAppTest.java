package eBayMain;

import java.net.MalformedURLException;

import org.openqa.selenium.WebElement;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import resources.ExcelUtils;
import testPages.CartPage;
import testPages.HomePage;
import testPages.LoginPage;
import testPages.ProductDetailsPage;
import wrappersTest.Automation_Utilities;

public class MainAppTest extends Automation_Utilities{

	
	private static AndroidDriver<AndroidElement> driver;
	
	String sUsername = "******@gmail.com";
	String sPassword = "*************";

	@BeforeTest
	public void beforeTest() throws MalformedURLException{
		fetchObjects();		
		LanchApp();
	}

	@Test(priority = 0)
	public void userLoginToApp() 
	{	
		new LoginPage(driver).clickSigninButton().checkLoginRadioButton()
		.enterUsername(sUsername)
		.clickOnContinueButton()
		.enterPassword(sPassword)
		.clickOnLoginButton();

	}

	@Test(dependsOnMethods = { "userLoginToApp" }, dataProvider="testData")
	public void searchProductInHomePage(String sNo, String sProduct)
	{

		new HomePage(driver)
		.searchProduct(sProduct).clickOnsearchedName()
		.selectProductFromList();
	}

	@Test(dependsOnMethods = { "searchProductInHomePage" })
	public void getProductDetailsFromPDP()
	{
		new ProductDetailsPage(driver)
		.verifyProductNameInPDP()
		.verifyProductDescriptionInPDP()
		.verifyProductPriceInPDP()
		.clickAddToCartButton()
		.clickOnCartIcon();
	}

	@Test(dependsOnMethods = { "getProductDetailsFromPDP" })
	public void getProductDetailsFromCartPage()
	{
		new CartPage(driver)
		.verifyProductNameInCartPage()
		.verifyProductDescInCartPage()
		.verifyProductPriceInCartPage();
	}

	@Test(dependsOnMethods = { "getProductDetailsFromCartPage" })
	public void validateProductDetialsInCartPage()
	{
		compareValues(productNameInCartPage,productNameInPDP);
		compareValues(productDescriptionInCartPage,productDescriptionInPDP);
		compareValues(productPriceInCartPage,productPriceInPDP);
	
	}

	@AfterTest
	public void quit()
	{
		closeBrowser();
	}

	@DataProvider(name = "testData")
	public Object[][] testData()
	{    	
		ExcelUtils eu = new ExcelUtils();
		Object data[][] = eu.readData();
		return data;

	}   

}
