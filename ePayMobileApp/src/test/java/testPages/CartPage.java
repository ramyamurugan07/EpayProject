package testPages;

import org.openqa.selenium.WebElement;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import wrappersTest.Automation_Utilities;

public class CartPage extends Automation_Utilities{

	public CartPage(AndroidDriver<AndroidElement> driver){
		this.driver = driver;
	}
	public CartPage verifyProductNameInCartPage(){
		productNameInCartPage = verifyAndGetByPath("CartPage.ProductName.Xpath");
		return this;
	}
	public CartPage verifyProductDescInCartPage(){
		productDescriptionInCartPage = verifyAndGetByPath("CartPage.ProductDesc.Xpath");
		return this;
	}
	public CartPage verifyProductPriceInCartPage(){
		productPriceInCartPage = verifyAndGetByPath("CartPage.ProductPrice.Xpath");
		return this;
	}
}
