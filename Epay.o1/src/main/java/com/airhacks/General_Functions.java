package com.airhacks;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import javax.swing.JDialog;
import javax.swing.JOptionPane;

import org.apache.poi.hpsf.Date;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openqa.selenium.WebDriver;

public class General_Functions {
	public static void main(String args[]) {
		System.out.println("Generic functions");
	}
	Excel_Utilities Obj_Excel_Utilities = new Excel_Utilities();
	Main Obj_Main = new Main();
	public void makeFolder(String sFilePath){

		File file = new File(sFilePath);
		file.mkdir();
	}
	public void Message(String sMessage, String sTitleBar, long lSleep) throws InterruptedException {

		JOptionPane pane = new JOptionPane(sMessage, JOptionPane.INFORMATION_MESSAGE);
		JDialog dialog = pane.createDialog(null, sTitleBar);
		dialog.setModal(false);
		dialog.setVisible(true);
		Thread.sleep(lSleep);
		dialog.setVisible(false);

	}

	public String StringReverse(String sFileName) {

		String sReverseFileName = sFileName;
		sFileName = "";
		for (int j = sReverseFileName.length(); j > 0; --j) {
			sFileName = sFileName + (sReverseFileName.charAt(j - 1));
		}
		return sFileName;

	}

	public String SubString(String sFullString, int iStart, int iEnd) {

		String sSubString = sFullString;
		sFullString = "";
		for (int j = iStart - 1; j < iEnd; j++) {
			if (j < sSubString.length()) {

				sFullString = sFullString + (sSubString.charAt(j));

			}
		}
		return sFullString;

	}
	
	
	
		
	}

