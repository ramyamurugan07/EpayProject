package testPages;

import org.openqa.selenium.WebElement;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import wrappersTest.Automation_Utilities;

public class HomePage extends Automation_Utilities{
	
	public HomePage(AndroidDriver<AndroidElement> driver){
		this.driver = driver;
	}
	public HomePage searchProduct(String product){
		typeByID("Home.SearchBar.Id",product+"\n");
		return this;		
	}
	public HomePage clickOnsearchedName(){
		clickByXPath("Home.ClickProduct.Xpath");
		return this;		
	}
	public ProductDetailsPage selectProductFromList(){
		clickByXPath("Home.Product.Xpath");
		return new ProductDetailsPage(driver);
	}
}
