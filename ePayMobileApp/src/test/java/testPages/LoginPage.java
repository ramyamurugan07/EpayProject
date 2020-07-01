package testPages;

import org.openqa.selenium.WebElement;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import wrappersTest.Automation_Utilities;

public class LoginPage extends Automation_Utilities{

	public LoginPage(AndroidDriver<AndroidElement> driver){
		this.driver = driver;
	}
	public LoginPage clickSigninButton(){
		clickByID("Login.SigninButton.Id");
		return this;
	}
	public LoginPage checkLoginRadioButton(){
		verifyByID("Login.LoginRadioButton.Id");
		return this;		
	}
	public LoginPage enterUsername(String userName){
		typeByID("Login.UserName.Id", userName);
		return this;
	}
	public LoginPage clickOnContinueButton(){
		clickByXPath("Login.Continue.Xpath");
		return this;
	}	
	public LoginPage enterPassword(String passWord){
		typeByID("Login.Password.Id", passWord);
		takeScreenshot();
		return this;
	}
	public HomePage clickOnLoginButton(){
		clickByID("Login.LoginButton.Id");
		return new HomePage(driver);
	}
}
