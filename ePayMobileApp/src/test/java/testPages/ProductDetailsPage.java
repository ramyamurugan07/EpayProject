package testPages;

import org.openqa.selenium.WebElement;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import wrappersTest.Automation_Utilities;

public class ProductDetailsPage extends Automation_Utilities {


	public ProductDetailsPage(AndroidDriver<AndroidElement> driver){
		this.driver = driver;
	}
	public ProductDetailsPage verifyProductNameInPDP(){
		productNameInPDP = verifyAndGetByID("PDP.ProductName.Id");
		return this;
	}
	public ProductDetailsPage verifyProductDescriptionInPDP(){
		productDescriptionInPDP=verifyAndGetByID("PDP.ProductDesc.Id");
		return this;
	}
	public ProductDetailsPage verifyProductPriceInPDP(){
		productPriceInPDP=verifyAndGetByID("PDP.ProductPrice.Id");
		return this;
	}
	public ProductDetailsPage clickAddToCartButton(){
		clickByID("PDP.AddtoCart.Id");
		return this;
	}
	public CartPage clickOnCartIcon(){
		clickByAccessibilityID("PDP.Cart.AccId");
		return new CartPage(driver);
	}
}
