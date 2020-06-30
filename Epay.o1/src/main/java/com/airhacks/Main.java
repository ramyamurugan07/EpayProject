package com.airhacks;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;

public class Main {
	static String sOutputLogfile;
	static String sTestcasePrefix = "DRV_";
	public static void main(String[] args) throws IOException, InterruptedException, AWTException

	{
		String sTestCase = null;
		String sFileName = "/Input.xlsm";
		String sReportDir = "/Output_Reports";
		String sWorkIndicator = "Execute";
		String sOutputFile = null ;
		//String sWorkIndicator = "Save_All_WS_TestCase";
		boolean bSequence_Flag = true;
		String sProtocol = null;
		FileInputStream File = null;
		XSSFWorkbook WorkBook = null, Output_Workbook = null;
		XSSFSheet TestCasesSheet = null, Sheet = null, Output_Sheet = null, Config_Sheet = null,sDriveSheet=null,Status_sheet=null;
		String sMailFlag,sOutputFile_Flag; 
		ArrayList<String> pAttachments = new ArrayList<String>();
		
		String currentDirectory = System.getProperty("user.dir");
		
		Excel_Utilities Obj_Excel_Utilities = new Excel_Utilities();
		Automation_Utilities Obj_AutomationUtilities = new Automation_Utilities();
		General_Functions Obj_General_Functions = new General_Functions();
		Main Obj_Main = new Main();
		
		File = Obj_Excel_Utilities.GetExcel(sFileName);
		WorkBook = Obj_Excel_Utilities.GetWorkbook(File);

		sDriveSheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, "Drive");
		Config_Sheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, "Config");
		
		String sAppName = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 23, 16);
		String sRepository = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 9, 2);
		

		sMailFlag = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 5, 2);
		sOutputFile_Flag = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 4, 2);
		
		
		String sFromMail = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 12, 16);
		String sAuthentication = Obj_Excel_Utilities.GetCellStringValue(Config_Sheet, 13, 16);
		
		Obj_General_Functions.makeFolder("Logs");
		Obj_General_Functions.makeFolder("Output_Reports");
		Obj_General_Functions.makeFolder("Screenshots");
		Obj_General_Functions.makeFolder("Test_Cases");
		
		DateFormat dateFormat = new SimpleDateFormat("ddMMMyyyy_HHmmss");
		Date Date = new Date();
		sOutputLogfile = System.getProperty("user.dir")+"/Logs/Log_"+dateFormat.format(Date)+ ".txt";
		Obj_Main.createFile(sOutputLogfile);		

		String sOutputFile_Squence = sReportDir+"/Output"+dateFormat.format(Date)+ ".xlsx";
		Obj_Excel_Utilities.CreateExcel(sOutputFile_Squence,"Output","","","");
		
		FileInputStream OutputFile_Sequence = Obj_Excel_Utilities.GetExcel(sOutputFile_Squence);                     
		XSSFWorkbook OutputWorkbook_Sequence = Obj_Excel_Utilities.GetWorkbook(OutputFile_Sequence);
		XSSFSheet OutputSheet_Sequence = Obj_Excel_Utilities.GetWorksheet_BySheetName(OutputWorkbook_Sequence, "Output");
		Obj_Excel_Utilities.createSequenceReport(sFileName, sOutputFile_Squence,OutputWorkbook_Sequence,OutputSheet_Sequence);
		

		String sAllTestCases = "";
		
		String sSubject = "";
		String sSequenceDriver, sTestCaseStatus = null;
		TestCasesSheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, "TestCases");
		sSequenceDriver = Obj_Excel_Utilities.GetCellStringValue(TestCasesSheet, 0, 25);
		
		int iExecutionFailureCount = 0;
		int iTestCasesRow ;
		String sTemp1 = null;
		
		if(sWorkIndicator.equals("Execute")){
			
			AndroidDriver<AndroidElement> Driver = Obj_AutomationUtilities.LaunchApp(sAppName);
			System.out.println(Driver);
			int iROW = -1,  iMaxRow = 0;
			int iRow = 0;
			Date dStart_TestCase = null;
			String sScreenshots = null;
			for (iTestCasesRow = 11; iTestCasesRow < 300; iTestCasesRow++)

			{
				
				String sAction;
				boolean bStatusMarked_Flag;
				sTestCaseStatus = "PASSED";
					sTestCase = Obj_Excel_Utilities.GetCellStringValue(TestCasesSheet, iTestCasesRow, 23);
					sProtocol = Obj_Excel_Utilities.GetCellStringValue(TestCasesSheet, iTestCasesRow, 22);
					sOutputFile = sReportDir +sFileName;
					iExecutionFailureCount = 0;
				Obj_Main.trace("sProtocol"+sProtocol+sTestCase , "BOTH");
				if (sTestCase.length() > 0)

				{				
					if (sProtocol.equals("Drive")) {
						
						
							iRow = 0;
						
							sSubject = sSubject.length() > 0 ? sSubject + " - " + sTestCase
									: "Status report for " + sTestCase;
							Sheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, "Drive");
							

							if (sOutputFile_Flag.equalsIgnoreCase("Yes"))
							{
								
								if(bSequence_Flag==true)
								{
								sOutputFile = sOutputFile_Squence;// Workbook name
								OutputFile_Sequence = Obj_Excel_Utilities.GetExcel(sOutputFile);  
								Output_Workbook = OutputWorkbook_Sequence;
								Obj_Excel_Utilities.AddExcelSheet(Output_Workbook, sOutputFile, sTestcasePrefix+sTestCase);
								Output_Sheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(OutputWorkbook_Sequence, sTestcasePrefix+sTestCase+"_"+Output_Workbook.getNumberOfSheets());
								Obj_Excel_Utilities.WriteExcel(Output_Workbook, OutputSheet_Sequence, sOutputFile, iTestCasesRow-10, 5, sTestcasePrefix+sTestCase+"_"+Output_Workbook.getNumberOfSheets());
								}
								
								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 0,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 0));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 1,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 1));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 2,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 2));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 3,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 3));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 4,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 4));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 5,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 5));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 6,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 6));

								Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 0, 7,

										Obj_Excel_Utilities.GetCellStringValue(Sheet, 0, 7));

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 0);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 1);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 2);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 3);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 4);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 5);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6);

								Obj_Excel_Utilities.COL_Header_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7);

							//	Obj_Excel_Utilities.ReadTestCase(sRepository, sProtocol,System.getProperty("user.dir") + "/Test_Cases/", sWorksheet,Output_Workbook, Output_Sheet,sOutputFile, "#^*^#");
								Obj_Excel_Utilities.ReadTestCase(sRepository, sProtocol,System.getProperty("user.dir") + "/Test_Cases/", sTestCase,Output_Workbook, Output_Sheet,sOutputFile, "#^*^#");
                                iMaxRow = Integer.parseInt(Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, 1, 26));
								
								dStart_TestCase = Obj_Main.StartTimer();
							}
							// end
						}
						
						try
						{
							
							for (iRow = 1; iRow < 899; iRow++)
							{
															
								sAction = Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 1);
								bStatusMarked_Flag = false;
								if (sAction.length() == 0)
								{
									break;
								}
								switch (sAction) {

								case "Click by XPath":
									Obj_AutomationUtilities.Click_By_XPath(Driver,Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3));
									break;
								case "Click by ID":
									Obj_AutomationUtilities.Click_By_ID(Driver,Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3));
									break;
								case "Click by Text":
									Obj_AutomationUtilities.Click_By_Text(Driver,Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3));
									break;

								case "Type by Xpath":
									Obj_AutomationUtilities.Type_By_xpath(Driver,Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3),Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4));
									break;
								case "Type By ID":
									Obj_AutomationUtilities.Type_By_ID(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3),
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4));
									break;
								case "Click by Classname":
									Obj_AutomationUtilities.Click_By_ClassName(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3));
									break;
								case "Type by Classname":
									Obj_AutomationUtilities.Click_By_ClassName(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3));
									break;
								case "Get Text":
									String sValue = Obj_AutomationUtilities.get_text(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3),Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4));
									
									break;
								case "Click by Index":
									Obj_AutomationUtilities.Click_by_Index(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3),Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4));
									
									break;
									
								case "Validation":
									Boolean bStatus = Obj_AutomationUtilities.validation(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3),Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4));
									if(bStatus==true){
										
										Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6, "PASSED");
										Obj_Excel_Utilities.PASSED_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6);
										Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4)+ "is displaying");
										Obj_Excel_Utilities.PASSED_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7);
										bStatusMarked_Flag = true;
									}
								
									else if (bStatus==false){
										
										Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6, "FAILED");
										Obj_Excel_Utilities.FAILED_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6);
										Obj_Excel_Utilities.WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 4)+"is not displaying");
										Obj_Excel_Utilities.FAILED_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7);
										bStatusMarked_Flag = true;
									}
									break;
								case "ScrollTo":
									Obj_AutomationUtilities.Scroll_To(Driver,
											Obj_Excel_Utilities.GetCellStringValue(Output_Sheet, iRow, 3)
											);
									break;
								default:
									Obj_Main.trace("Exited on Row " + iRow + " (Action=" + sAction + ")", "BOTH");
									break;

								}

								if (bStatusMarked_Flag == false ) {
									Obj_Excel_Utilities.Mark_Passed(sOutputFile_Flag, Output_Workbook, Output_Sheet, sOutputFile, iRow,"");
								}
							}
							
							
						}
						catch (Exception e)

						{
							
							sTestCaseStatus = "*#PWF_";
							//sTestCaseStatus = "FAILED";
							if (sOutputFile_Flag.equalsIgnoreCase("Yes"))

							{
								String sError_Message;
								String sScreenshotFileName;
								sTemp1 = String.format("_Row_%d", iRow);
								sScreenshotFileName = Obj_General_Functions.StringReverse(sOutputFile);
								sScreenshotFileName = Obj_General_Functions.SubString(sScreenshotFileName, 6,
										sScreenshotFileName.indexOf("/", 0));
								sScreenshotFileName = Obj_General_Functions.StringReverse(sScreenshotFileName) + sTemp1 + ".jpg";
								if(Driver != null){
									Obj_AutomationUtilities.TakeScreenshot(Driver, sScreenshots, sScreenshotFileName);
								}
								pAttachments.add(sScreenshots+"/"+sScreenshotFileName);
								
									sError_Message = "##_EXCEPTION_## - " + e.getMessage() + "\n" + sReportDir + "/"
											+ sScreenshotFileName;
								
								Obj_Excel_Utilities.Mark_Failed(sOutputFile_Flag, Output_Workbook, Output_Sheet, sOutputFile, iRow,
										sError_Message);
							}
							Obj_Main.trace("Exception is found here: " + e, "BOTH");
						}			
					}
					Obj_Main.trace("Testcase :"+sTestCase +" finished....", "BOTH");

					String sTime_TestCase = Obj_Main.StopTimer(dStart_TestCase);
					sAllTestCases = sAllTestCases+ sTestCase + "\t\t\t- " + sTestCaseStatus + "\t\t\t- Time Taken: " + sTime_TestCase + "\n" ;
					//writing to autodrive sequence excel
					if((sSequenceDriver.trim().isEmpty()) || (sSequenceDriver.trim().equalsIgnoreCase("SEQUENCE")) || (sSequenceDriver.trim().equalsIgnoreCase("SELECTION")))

					{
					if(sTestCaseStatus == "*#PWF_")
					{
						Obj_Excel_Utilities.Mark_Sequence(sOutputFile_Flag, OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iTestCasesRow-10,3, "PASSED_WITH_FAILURES", "ORANGE");

					}
					else if(sTestCaseStatus == "FAILED")
					{
						Obj_Excel_Utilities.Mark_Sequence(sOutputFile_Flag, OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iTestCasesRow-10,3, "FAILED", "RED");
					}
					else
					{
						Obj_Excel_Utilities.Mark_Sequence(sOutputFile_Flag, OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iTestCasesRow-10,3, "PASSED", "GREEN");
					}
					
					Obj_Excel_Utilities.WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iTestCasesRow-10, 4,sTime_TestCase);
					Obj_Excel_Utilities.setColor(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iTestCasesRow-10, 4, "VIOLET");
					}
					
				}

			}
			
			
				

		else if(sWorkIndicator.equals("Save_TestCase"))
		{
			
			String sDirectory = currentDirectory+"/Test_Cases/";
			String sDelimeter = "#^*^#";
			Obj_Excel_Utilities.SaveTestCase(sRepository, sDirectory, Obj_Excel_Utilities.GetCellStringValue(sDriveSheet, 3, 9), sDriveSheet, sDelimeter);
			Obj_Main.trace("Main :: TestCase - "+Obj_Excel_Utilities.GetCellStringValue(sDriveSheet, 3, 9)+" saved successfully", "BOTH");

		}
		else if(sWorkIndicator.equals("Save_All_WS_TestCase"))
		{
			String sDirectory = currentDirectory+"/Test_Cases/";
			XSSFSheet Worksheet = null;
			String sDelimeter = "#^*^#";
			Sheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, "TestCases");
			for(int iRow=2;iRow<300;iRow++)
			{                                              
				sTestCase = Obj_Excel_Utilities.GetCellStringValue(Sheet, iRow, 28);
				
				if(sTestCase.length()>0)
				{
					Worksheet = Obj_Excel_Utilities.GetWorksheet_BySheetName(WorkBook, sTestCase);
					Obj_Excel_Utilities.SaveTestCase(sRepository, sDirectory, sTestCase, Worksheet, sDelimeter);
				}
				else
				{
					break;
				}
			}
		}
		Obj_Main.trace("Execution completed", "BOTH");
		
	}

	public Date StartTimer() {
		return new Date();
	}

	public String StopTimer(Date dStartDateTime) {

		Date dStopDateTime = new Date();
		
		if(dStartDateTime != null){
			long difference = dStopDateTime.getTime() - dStartDateTime.getTime();
			long seconds = TimeUnit.MILLISECONDS.toSeconds(difference);
			double minutes = seconds / 60.0; 
			return seconds + " Secs (" + String.format("%.2f",minutes) + " Mins)";

		}
		return   "0.00 Secs (0 Mins)";
	}

	public void trace(String sText,String sWorkIndicator){

		FileWriter fw = null;
		try {
			File file = new File(Main.sOutputLogfile);
			fw = new FileWriter(file, true);
			Date date = new Date();
			DateFormat dateFormat = new SimpleDateFormat("dd-MMM-yy hh:mm");
			sText = dateFormat.format(date) + " - " +sText +"\n";
			if(sWorkIndicator.equals("CONSOLE")){
				System.out.println(sText);
			}else if(sWorkIndicator.equals("LOG")){
				fw.append(sText);
			}else if(sWorkIndicator.equals("BOTH")){
				System.out.println(sText);
				fw.append(sText);
			}else if(sWorkIndicator.equals("Creation_Failed") || sWorkIndicator.equals("Creation_Successful")){
				sText = "\n*****************************\n"+dateFormat.format(date)+"\n*****************************\n\n";
				fw.append(sText);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
		finally {
			try {
				fw.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	public void createFile(String filepath){

		try {

			File file = new File(filepath);
			file.createNewFile();
			trace("file is created","Creation_Successful");

		}
		catch (IOException e) {

			e.printStackTrace();

			trace("createFile() Finished","Creation_Failed");
		}

		trace("createFile() Finished","LOG");
	}
}
