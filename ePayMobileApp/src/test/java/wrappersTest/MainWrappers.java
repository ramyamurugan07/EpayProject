package wrappersTest;

import java.net.MalformedURLException;

public interface MainWrappers {
	public void LanchApp() throws MalformedURLException;
	public void typeByID(String sID, String svalue);
	public void clickByID(String sID);
	public void clickByXPath(String sXpath);
	public void verifyByID(String sID);
	public String verifyAndGetByID(String sID);
	public void clickByAccessibilityID(String sID);
	public String verifyAndGetByPath(String sxpath);
	public void compareValues(String sActual, String sExpected);
}
