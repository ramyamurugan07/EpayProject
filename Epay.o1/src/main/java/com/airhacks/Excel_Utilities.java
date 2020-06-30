package com.airhacks;

import java.awt.AWTException;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.text.SimpleDateFormat;
import java.util.Scanner;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Utilities {
	public static void main(String args[]) {
		System.out.println("Excel Utilities");
	}
	Main Obj_Main = new Main();
public FileInputStream GetExcel(String sFilePathName) throws FileNotFoundException {
		
		File file = new File(System.getProperty("user.dir") + sFilePathName);
		
		FileInputStream fs = new FileInputStream(file);
		return fs;
	}

	// Reading the Excel & returning XSSFWorkbook variable pointing to Workbook
	public XSSFWorkbook GetWorkbook(FileInputStream fs) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		return workbook;
	}

	// Reading the Excel & returning XSSFSheet variable pointing to respective Excel
	// Worksheet Index
	public XSSFSheet GetWorksheet_BySheetNumber(XSSFWorkbook workbook, Integer iIndex) throws IOException {
		XSSFSheet sheet = workbook.getSheetAt(iIndex);
		return sheet;
	}

	// Reading the Excel & returning XSSFSheet variable pointing to respective Excel
	// Worksheet Name
	public XSSFSheet GetWorksheet_BySheetName(XSSFWorkbook workbook, String sName) throws IOException {
		XSSFSheet sheet = workbook.getSheet(sName);
		return sheet;
	}

	// To Read Excel Cell Value & return String as Cell Value
	public String GetCellStringValue(XSSFSheet sheet, Integer iRow, Integer iCol) {
		Row Row;
		Cell cell;
		String strCellValue = "";

		Row = sheet.getRow(iRow);

		if (Row != null) {
			cell = Row.getCell(iCol);
			if (cell != null) {
				// System.out.println(cell);
				// strCellValue = null;
				if (cell != null) {
					// System.out.println(cell.getCellType());
					switch (cell.getCellType()) {

					case STRING:
						strCellValue = cell.toString();
						break;
					case NUMERIC:

						if (DateUtil.isCellDateFormatted(cell)) {
							SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
							strCellValue = dateFormat.format(cell.getDateCellValue());
						} else {
							int iValue = (int) cell.getNumericCellValue();
							strCellValue = String.valueOf(iValue);
							// TODO:: need to modify 
							// strCellValue = String.valueOf(cell.getNumericCellValue());
						}
						break;
					case BOOLEAN:
						strCellValue = String.valueOf(cell.getBooleanCellValue());
						break;
					case BLANK:
						strCellValue = "";
						break;
					case FORMULA:
						strCellValue = cell.getRichStringCellValue().getString();
					}
				}
			}
		}
		strCellValue.trim();
		// System.out.println("Excel Utilities");
		return strCellValue;
	}


	public void CreateExcel(String sFileNameWithPath, String sSheetName1, String sSheetName2, String sSheetName3,
			String sSheetName4) throws IOException {
		Workbook wb = new XSSFWorkbook();
		File file = new File(System.getProperty("user.dir") + sFileNameWithPath);
		FileOutputStream fileOut = new FileOutputStream(file);
		sSheetName1.trim();
		sSheetName2.trim();
		sSheetName3.trim();
		sSheetName4.trim();
		if (sSheetName1.length() > 0) {
			org.apache.poi.ss.usermodel.Sheet sheet1 = wb.createSheet(sSheetName1);
			for (int iRow = 0; iRow < 1000; iRow++) {
				sheet1.createRow(iRow);
			}
		}
		if (sSheetName2.length() > 0) {
			org.apache.poi.ss.usermodel.Sheet sheet2 = wb.createSheet(sSheetName2);
			for (int iRow = 0; iRow < 1000; iRow++) {
				sheet2.createRow(iRow);
			}
		}
		if (sSheetName3.length() > 0) {
			org.apache.poi.ss.usermodel.Sheet sheet3 = wb.createSheet(sSheetName3);
			for (int iRow = 0; iRow < 1000; iRow++) {
				sheet3.createRow(iRow);
			}
		}
		if (sSheetName4.length() > 0) {
			org.apache.poi.ss.usermodel.Sheet sheet4 = wb.createSheet(sSheetName4);
			for (int iRow = 0; iRow < 1000; iRow++) {
				sheet4.createRow(iRow);
			}
		}
		wb.write(fileOut);
		fileOut.close();
	}

	public void WriteExcel(XSSFWorkbook pWorkBook, XSSFSheet sheet, String sFileNameWithPath, Integer iRow,
			Integer iCol, String sText) throws IOException {
		Row Row;
		Cell Cell;
		Row = sheet.getRow(iRow);
		if (Row == null) {
			sheet.createRow(iRow);
		}
		Cell = Row.getCell(iCol);
		if (Cell == null) {
			Row.createCell(iCol);
		}
		CellStyle style = pWorkBook.createCellStyle();
		sheet.getRow(iRow).getCell(iCol).setCellValue(sText);
		if (iCol != 7) {
			sheet.autoSizeColumn(iCol);
		} else {
			sheet.setColumnWidth(iCol, 10000);
		}
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		XSSFFont font = pWorkBook.createFont();
		((org.apache.poi.ss.usermodel.Font) font).setFontHeight((short) 180);
		style.setFont((org.apache.poi.ss.usermodel.Font) font);
		sheet.getRow(iRow).getCell(iCol).setCellStyle(style);
		File file = new File(System.getProperty("user.dir") + sFileNameWithPath);
		FileOutputStream out = new FileOutputStream(file);
		pWorkBook.write(out);
		out.close();
	}

	public void PASSED_Color(XSSFWorkbook pWorkBook, XSSFSheet sheet, String sFileNameWithPath, Integer iRow,
			Integer iCol) throws IOException {
		Row Row;
		Cell Cell;
		Row = sheet.getRow(iRow);
		if (Row != null) {
			Cell = Row.getCell(iCol);
			if (Cell != null) {
				CellStyle style = pWorkBook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
				XSSFFont font = pWorkBook.createFont();
				((org.apache.poi.ss.usermodel.Font) font).setColor(IndexedColors.WHITE.getIndex());
				((org.apache.poi.ss.usermodel.Font) font).setBold(true);
				((org.apache.poi.ss.usermodel.Font) font).setFontHeight((short) 180);
				style.setFont((org.apache.poi.ss.usermodel.Font) font);
				sheet.getRow(iRow).getCell(iCol).setCellStyle(style);
				File file = new File(System.getProperty("user.dir") + sFileNameWithPath);
				FileOutputStream out = new FileOutputStream(file);
				pWorkBook.write(out);
				out.close();
			}
		}
	}

	public void FAILED_Color(XSSFWorkbook pWorkBook, XSSFSheet sheet, String sFileNameWithPath, Integer iRow,
			Integer iCol) throws IOException {
		Row Row;
		Cell Cell;
		Row = sheet.getRow(iRow);
		if (Row != null) {
			Cell = Row.getCell(iCol);
			if (Cell != null) {
				CellStyle style = pWorkBook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.RED.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
				XSSFFont font = pWorkBook.createFont();
				((org.apache.poi.ss.usermodel.Font) font).setColor(IndexedColors.WHITE.getIndex());
				((org.apache.poi.ss.usermodel.Font) font).setBold(true);
				((org.apache.poi.ss.usermodel.Font) font).setFontHeight((short) 180);
				style.setFont((org.apache.poi.ss.usermodel.Font) font);
				sheet.getRow(iRow).getCell(iCol).setCellStyle(style);
				File file = new File(System.getProperty("user.dir") + sFileNameWithPath);
				FileOutputStream out = new FileOutputStream(file);
				pWorkBook.write(out);
				out.close();
			}
		}
	}

	public void COL_Header_Color(XSSFWorkbook pWorkBook, XSSFSheet sheet, String sFileNameWithPath, Integer iRow,
			Integer iCol) throws IOException {
		Row Row;
		Cell Cell;
		Row = sheet.getRow(iRow);
		if (Row != null) {
			Cell = Row.getCell(iCol);
			if (Cell != null) {
				CellStyle style = pWorkBook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				XSSFFont font = pWorkBook.createFont();
				((org.apache.poi.ss.usermodel.Font) font).setColor(IndexedColors.WHITE.getIndex());
				((org.apache.poi.ss.usermodel.Font) font).setBold(true);
				((org.apache.poi.ss.usermodel.Font) font).setFontHeight((short) 180);
				style.setFont((org.apache.poi.ss.usermodel.Font) font);
				sheet.getRow(iRow).getCell(iCol).setCellStyle(style);
				File file = new File(System.getProperty("user.dir") + sFileNameWithPath);
				FileOutputStream out = new FileOutputStream(file);

				// FileOutputStream out = new FileOutputStream(new File(sFileNameWithPath));
				pWorkBook.write(out);
				out.close();
			}
		}
	}
	

	public void setColor(XSSFWorkbook pWorkBook, XSSFSheet sheet, String sFileNameWithPath, Integer iRow,
			Integer iCol, String bgColor) throws IOException {
		
		
		Row Row;
		Cell Cell;
		Row = sheet.getRow(iRow);
		if (Row != null) {
			Cell = Row.getCell(iCol);
			if (Cell != null) {
				CellStyle style = pWorkBook.createCellStyle();
				
				switch(bgColor){

				case "ORANGE":
					style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
					break;
					
				case "AQUA":
					style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
					break;
					
				case "GREEN":
					style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
					break;
				case "RED":
					style.setFillForegroundColor(IndexedColors.RED.getIndex());
					break;
					
				case "BLUE":
					style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
					break;
					
				case "VIOLET":
					style.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
					break;
					
				case "GOLD":
					style.setFillForegroundColor(IndexedColors.GOLD.getIndex());
					break;
					
				default:
					style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
					break;
					
				
				}
				
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setBorderTop(BorderStyle.THIN);
				style.setBorderBottom(BorderStyle.THIN);
				style.setBorderLeft(BorderStyle.THIN);
				style.setBorderRight(BorderStyle.THIN);
				XSSFFont font = pWorkBook.createFont();
				
				((org.apache.poi.ss.usermodel.Font) font).setColor(IndexedColors.WHITE.getIndex());
				((org.apache.poi.ss.usermodel.Font) font).setBold(true);
				((org.apache.poi.ss.usermodel.Font) font).setFontHeight((short) 180);
				style.setFont((org.apache.poi.ss.usermodel.Font) font);
				sheet.getRow(iRow).getCell(iCol).setCellStyle(style);
				File file = new File(System.getProperty("user.dir") + sFileNameWithPath);
				FileOutputStream out = new FileOutputStream(file);
				pWorkBook.write(out);
				out.close();
			}
		}
	}
	@SuppressWarnings("static-access")
	public void CreateOutputFile(XSSFWorkbook WorkBook, XSSFSheet Sheet, int iRow, int iCol,
			XSSFWorkbook Output_Workbook, XSSFSheet Output_Sheet, String sOutputFile) throws IOException {
		String sAction;
		Excel_Utilities Obj_Excel_Utilities = new Excel_Utilities();
		for (iRow = 2; iRow < 899; iRow++) {
			sAction = Obj_Excel_Utilities.GetCellStringValue(Sheet, iRow, iCol);
			if (sAction.length() > 0) {
				WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, iCol,
						Obj_Excel_Utilities.GetCellStringValue(Sheet, iRow, iCol));
			} else {
				if (Obj_Excel_Utilities.GetCellStringValue(Sheet, iRow, 0).length() == 0) {
					return;
				} else {
					WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, iCol, "");
				}
			}
		}
	}

	
	
	public void createSequenceReport(String sFileName, String sOutputFile_Squence, XSSFWorkbook OutputWorkbook_Sequence,
			XSSFSheet OutputSheet_Sequence) throws IOException {
			
			FileInputStream File = GetExcel(sFileName);
			XSSFWorkbook WorkBook = GetWorkbook(File);
			XSSFSheet Sheet = GetWorksheet_BySheetName(WorkBook, "TestCases");

			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 0,
					GetCellStringValue(Sheet, 10, 21));
			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 1,
					GetCellStringValue(Sheet, 10, 22));
			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 2,
					GetCellStringValue(Sheet, 10, 23));
			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 3,
					GetCellStringValue(Sheet, 10, 24));

			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 4, "Time Taken");
			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 5, "Report Link");

			COL_Header_Color(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 0);
			COL_Header_Color(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 1);
			COL_Header_Color(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 2);
			COL_Header_Color(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 3);
			COL_Header_Color(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 4);
			COL_Header_Color(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, 0, 5);

			for (int iRow = 11; iRow < 60; iRow++) {
				String sSequence = GetCellStringValue(Sheet, iRow, 23);

				if (sSequence.length() > 0) {
					WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow - 10, 0,
							String.valueOf(iRow - 10));
					// Excel_Utilities.WriteExcel(OutputWorkbook_Sequence,
					// OutputSheet_Sequence, sOutputFile_Squence, iRow-10, 0,
					// Excel_Utilities.GetCellStringValue(Sheet, iRow, 21));
					WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow - 10, 1,
							GetCellStringValue(Sheet, iRow, 22));
					WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow - 10, 2,
							GetCellStringValue(Sheet, iRow, 23));
					WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow - 10, 3,
							"<Not Executed>");
//					String sTestcase_Name = "DRV_" + GetCellStringValue(Sheet, iRow, 23)+"_"+(iRow-9);
//					setHyperLink(OutputWorkbook_Sequence,OutputSheet_Sequence,sTestcase_Name,iRow - 10,5);

				} else {
					break;
				}
			}
		}

	public void SaveTestCase(String sRepository, String sDirectory, String sTestCaseName, XSSFSheet Sheet,
			String sDelimeter) throws IOException {
		String sContent, sAction, sTestCaseFile;
		// Excel_Utilities Obj_Excel_Utilities = new Excel_Utilities();
		String sProtocol = GetCellStringValue(Sheet, 0, 53);

		if (sRepository.equals("Local")) {
			if (sProtocol.equals("DRIVE")) {
				sTestCaseFile = sDirectory + "Auto_Drive_" + sTestCaseName + ".txt";

				File pFile = new File(sTestCaseFile);
				FileOutputStream fos = new FileOutputStream(pFile);
				BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));

				for (int iRow = 1; iRow < 600; iRow++) {
					sContent = "";
					sAction = GetCellStringValue(Sheet, iRow, 1);
					if (sAction.length() > 0) {
						sContent = GetCellStringValue(Sheet, iRow, 1) + sDelimeter;
						sContent = sContent + GetCellStringValue(Sheet, iRow, 2) + sDelimeter;
						sContent = sContent + GetCellStringValue(Sheet, iRow, 3) + sDelimeter;
						sContent = sContent + GetCellStringValue(Sheet, iRow, 4) + sDelimeter;
						sContent = sContent + GetCellStringValue(Sheet, iRow, 5) + sDelimeter;
						bw.write(sContent);
						bw.newLine();
					} else {
						break;
					}
				}

				bw.close();
			}
		}
	}

	public void ReadTestCase(String sRepository,String sProtocol, String sDirectory, String sTestCaseName, XSSFWorkbook Output_Workbook, XSSFSheet Output_Sheet,String sOutputFile, String sDelimeter) throws IOException

    {

           String sContent,sAction,sTestCaseFile;

           Excel_Utilities Obj_Excel_Utilities = new Excel_Utilities();

           int iRow=1,iCount=0;

           String sData;

           int iIndex=0,iNextIndex=0;

          

           if(sRepository.equals("Local"))

           {

                  if(sProtocol.equals("Drive"))

                  {

                        sTestCaseFile =  sDirectory+ "Auto_Drive_"+sTestCaseName+".txt";



                        File pFile = new File(sTestCaseFile);

                        Scanner MyReader = new Scanner(pFile);

                       

                        while (MyReader.hasNextLine())

                        {

                               iCount=1;

                              

                               sData = MyReader.nextLine();



                               while(sData.length()>0)

                               {

                                      iNextIndex = sData.indexOf(sDelimeter);

                                      sContent = sData.substring(0, iNextIndex);

                                      sData = sData.substring(iNextIndex+sDelimeter.length(), sData.length());

                                      iIndex = iNextIndex;

                                     

                                      switch(iCount)

                                      {

                                             case 1:

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 0, String.valueOf(iRow));

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 1, sContent);

                                                    break;

                                             case 2:

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 2, sContent);

                                                    break;

                                             case 3:

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 3, sContent);

                                                    break;

                                             case 4:

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 4, sContent);

                                                    break;

                                             case 5:

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 5, sContent);

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6, "<Not Executed>");

                                                    WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, "");

                                                    break;

                                      }

                                      iCount++;                                      

                               }

                               iRow++;

                        }

                        WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, 1, 26, Integer.toString(iRow));

                        MyReader.close();

                  }


           }

    }

//////////////
	public void Mark_Sequence(String sOutputFile_Flag,XSSFWorkbook OutputWorkbook_Sequence,XSSFSheet OutputSheet_Sequence,String sOutputFile_Squence, int iRow, int iColumn,String sStatus, String sColor ) throws IOException{

		if (sOutputFile_Flag.equalsIgnoreCase("Yes")){

			WriteExcel(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow, iColumn, sStatus); 
			if(sStatus.equals("FAILED")){
				setColor(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow, iColumn, "RED");
			}else if(sStatus.equals("PASSED")){
				setColor(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow, iColumn, "GREEN");
			}else{
				setColor(OutputWorkbook_Sequence, OutputSheet_Sequence, sOutputFile_Squence, iRow, iColumn, sColor);
			}
		}

	}
	public void Mark_Passed(String sOutputFile_Flag, XSSFWorkbook Output_Workbook, XSSFSheet Output_Sheet,
			String sOutputFile, int iRow,String sTime) throws IOException {
		
		if (sOutputFile_Flag.equalsIgnoreCase("Yes"))

		{
			WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6, "PASSED");
			PASSED_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6);

			sTime = sTime.trim();
			if(sTime != null && sTime != ""){
				if(sTime.endsWith("is already stopped")){
					WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, sTime);
					setColor(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7,"ORANGE");
				}else{
					WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, sTime);
					setColor(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7,"AQUA");
				}
			}else{
				WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, sTime);
			}

		}

	}



	public void Mark_Failed(String sOutputFile_Flag, XSSFWorkbook Output_Workbook, XSSFSheet Output_Sheet,
			String sOutputFile, int iRow, String sError_Message) throws IOException, AWTException {

		

		if (sOutputFile_Flag.equalsIgnoreCase("Yes"))

		{
			WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6, "FAILED");
			FAILED_Color(Output_Workbook, Output_Sheet, sOutputFile, iRow, 6);

			sError_Message = sError_Message.trim();

			if(sError_Message != null && sError_Message != ""){

				WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, sError_Message);
				setColor(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, "RED");

			}else{
				WriteExcel(Output_Workbook, Output_Sheet, sOutputFile, iRow, 7, sError_Message);
			}

		}

	}
	public String AddExcelSheet(XSSFWorkbook pWorkBook, String sFileNameWithPath, String sSheetName1) throws IOException 
	{
	try
	{
	 sSheetName1.trim();
	 if (sSheetName1.length() > 0) 
	{                          
		 File file = new File(System.getProperty("user.dir") + sFileNameWithPath);//variable
	FileOutputStream out = new FileOutputStream(file);
	XSSFSheet Sheet1 = null;
	sSheetName1 = sSheetName1.format("%s_%d", sSheetName1,pWorkBook.getNumberOfSheets()+1);
	pWorkBook.createSheet(sSheetName1);
	Sheet1 = GetWorksheet_BySheetName(pWorkBook, sSheetName1);;
	for (int iRow = 0; iRow < 1000; iRow++) 
	{
	Sheet1.createRow(iRow);
	}
	pWorkBook.write(out);
	 out.close();
	}
	}
	catch (IOException e) 
	{                    
	e.printStackTrace();
	}
	return sSheetName1;
	}

	public void SetHyperlink(String FilePath, int iRow, int iCol, String sSourceSheetName,
			String sTargetSheetName, String sHyperlinkText, String sTargetRange) {
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		XSSFCell Cell;
		XSSFRow Row;
		FileInputStream fileIS;
		try {
			fileIS = new FileInputStream(System.getProperty("user.dir") +FilePath);
			
			// Get sheet access
			workbook = new XSSFWorkbook(fileIS);
			sheet = workbook.getSheet(sSourceSheetName);
			Row = sheet.getRow(iRow);
			if (Row == null) {
				sheet.createRow(iRow);
				Row = sheet.getRow(iRow);
			}
			Cell = Row.getCell(iCol);
			if (Cell == null) {
				Row.createCell(iCol);
				Cell = Row.getCell(iCol);
			}
			Cell.setCellValue(sHyperlinkText);
			CreationHelper createHelper = workbook.getCreationHelper();
			// Making hyperlinks blue and underlined
			CellStyle style = workbook.createCellStyle();
			Font hlink_font = workbook.createFont();
			hlink_font.setUnderline(Font.U_SINGLE);
			hlink_font.setColor(IndexedColors.BLUE.getIndex());
			style.setFont(hlink_font);
			// create a target sheet and cell
			
			XSSFHyperlink link2 = (XSSFHyperlink) createHelper.createHyperlink(HyperlinkType.DOCUMENT);
			link2.setAddress("'" + sTargetSheetName + "'!" + sTargetRange);
			Cell.setHyperlink(link2);
			Cell.setCellStyle(style);
			fileIS.close();
			// Close input stream
			// Write to file
			//FileOutputStream out = new FileOutputStream(FilePath);
			FileOutputStream out = new FileOutputStream(System.getProperty("user.dir") +FilePath);
			//FileOutputStream out = new FileOutputStream(System.getProperty("user.dir") + FilePath);
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
