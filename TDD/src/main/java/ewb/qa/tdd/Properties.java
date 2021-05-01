package ewb.qa.tdd;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Properties {
	public WebDriver driver;
	public ExcelObj excelObj;
	public SeleniumObj seleniumObj;
	public String strcustomernumber;
	public String strreferencefield;
	public String strPath;
	
	private static String browser;
	private static String baseurl;
	
	private static String ArchMaptestcaseid;
	private static String ArchMaptestcasedesciption;
	private static String ArchMapfunccode;
	private static int ArchMapiteration;
	private static String ArchMapexecutionflag;
	private static String ArchMapresult;
	private static int ArchMaptestcycle;
	private static int ArchMaptestrun;

	private static String DataSourcetestcaseid;
	private static String DataSourceuser;
	private static String DataSourcemainmodule;
	private static String DataSourcesubmodule;
	private static String DataSourcefunctioncode;
	private static String DataSourcefunctionality;
	private static String DataSourcemainmenu;
	private static String DataSourcesubmenu;
	
	private static String Funcmodule;
	private static String Funcfield;
	private static String Funcelementid;
	private static String Funcelementxpath;
	private static String Funclinktext;
	private static String Funcpageurl;
	private static String Functitle;
	private static String Funcelementtype;
	private static String Funcvalue;
	private static String Funcaction;
	private static String Funcmenureference;
	private static String Funcworksheetreference;
	private static String Funcfieldreference;
	private static String Funcmessage;
	private static String Funcresult;
	
	private static String Menumodule;
	private static String Menuworksheetcode_main;
	private static String Menuworksheetcode_sub;
	private static String Menuworksheet;
	
	private static String MainMenumodule;
	private static String MainMenufield;
	private static String MainMenuelementid;
	private static String MainMenuelementxpath;
	private static String MainMenulinktext;
	private static String MainMenupageurl;
	private static String MainMenutitle;
	private static String MainMenuelementtype;
	private static String MainMenuvalue;
	private static String MainMenuaction;
	private static String MainMenumenureference;
	private static String MainMenuworksheetreference;
	private static String MainMenufieldreference;
	private static String MainMenumessage;
	private static String MainMenuresult;
	
	private static String SubMenumodule;
	private static String SubMenufield;
	private static String SubMenuelementid;
	private static String SubMenuelementxpath;
	private static String SubMenulinktext;
	private static String SubMenupageurl;
	private static String SubMenutitle;
	private static String SubMenuelementtype;
	private static String SubMenuvalue;
	private static String SubMenuaction;
	private static String SubMenumenureference;
	private static String SubMenuworksheetreference;
	private static String SubMenufieldreference;
	private static String SubMenumessage;
	private static String SubMenuresult;
	
	private static ArrayList<String> arrTestCases;
	private static ArrayList<String> arrDataSource;
	private static ArrayList<String> arrTellerMenu;
	private static ArrayList<String> arrTellerMenuMain;
	private static ArrayList<String> arrTellerMenuSub;
	private static ArrayList<String> arrTDDFunc;
	private static ArrayList<String> arrScript;
	private static ArrayList<String> arrTempStorage;
	private static ArrayList<String> arrTestData;
	private static ArrayList<String> arrMapTD;
	
	private static String Scriptmodule;
	private static String Scriptfield;
	private static String Scriptelementid;
	private static String Scriptelementxpath;
	private static String Scriptlinktext;
	private static String Scriptpageurl;
	private static String Scripttitle;
	private static String Scriptelementtype;
	private static String Scriptvalue;
	private static String Scriptaction;
	private static String Scriptmenureference;
	private static String Scriptworksheetreference;
	private static String Scriptfieldreference;
	private static String Scriptmessage;
	private static String Scriptresult;
	
	private static String CurrentTestCaseId;
	
	private static String testdataField;
	private static String testdataGivenvalue;
	
	public static void setArrMapTestData(ArrayList<String> arrlistMapTD)
	{
		arrMapTD = arrlistMapTD;
	}
	
	public static ArrayList<String> getArrMapTestData()
	{
		return arrMapTD;
	}
	
	public static void setArrTestData(ArrayList<String> arrlistTestData)
	{
		arrTestData = arrlistTestData;
	}
	
	public static ArrayList<String> getArrTestData()
	{
		return arrTestData;
	}
	
	public static void setTestDataField(String givenValue)
	{
		testdataField = givenValue;
	}
	
	public static String getTestDataField()
	{
		return testdataField;
	}
	
	public static void TestData(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			testdataField = givenValue;
			break;
		
		case 1:
			testdataGivenvalue = givenValue;
			break;
		}
	}
	
	public static void setCurrentTestCaseId(String givenValue)
	{
		CurrentTestCaseId = givenValue;
	}
	
	public static String getCurrentTestCaseId()
	{
		return CurrentTestCaseId;
	}
	
	public static void setArrTempStorage(ArrayList<String> arrlistTempStorage)
	{
		arrTempStorage = arrlistTempStorage;
	}
	
	public static ArrayList<String> getArrTempStorage()
	{
		return arrTempStorage;
	}
	
	public static void setBrowser(String strBrowser)
	{
		browser = strBrowser;
	}
	
	public static String getBrowser()
	{
		return browser;
	}
	
	public static void setBaseUrl(String strBaseUrl)
	{
		baseurl = strBaseUrl;
	}
	
	public static String getBaseUrl()
	{
		return baseurl;
	}
	
	public static void setArrTestCases(ArrayList<String> arrlistTestCases)
	{
		arrTestCases = arrlistTestCases;
	}
	
	public static ArrayList<String> getArrTestCases()
	{
		return arrTestCases;
	}
	
	public static void setArrDataSource(ArrayList<String> arrlistDataSource)
	{
		arrDataSource = arrlistDataSource;
	}
	
	public static ArrayList<String> getArrDataSource()
	{
		return arrDataSource;
	}
	
	public static void setArrTellerMenu(ArrayList<String> arrlistTellerMenu)
	{
		arrTellerMenu = arrlistTellerMenu;
	}
	
	public static ArrayList<String> getArrTellerMenu()
	{
		return arrTellerMenu;
	}
	
	public static void setArrTellerMenuMain(ArrayList<String> arrlistTellerMenuMain)
	{
		arrTellerMenuMain = arrlistTellerMenuMain;
	}
	
	public static ArrayList<String> getArrTellerMenuMain()
	{
		return arrTellerMenuMain;
	}
	
	public static void setArrTellerMenuSub(ArrayList<String> arrlistTellerMenuSub)
	{
		arrTellerMenuSub = arrlistTellerMenuSub;
	}
	
	public static ArrayList<String> getArrTellerMenuSub()
	{
		return arrTellerMenuSub;
	}
	
	public static void setArrTDDFunc(ArrayList<String> arrlistTDDFunc)
	{
		arrTDDFunc = arrlistTDDFunc;
	}
	
	public static ArrayList<String> getArrTDDFunc()
	{
		return arrTDDFunc;
	}
	
	public static void setArrScript(ArrayList<String> arrlistScript)
	{
		//System.out.println("Array List from Properties \n" + arrlistScript);
		arrScript = arrlistScript;
	}
	
	public static ArrayList<String> getArrScript()
	{
		return arrScript;
	}
	
 	public void iWebDriverElementID(String strElement, String strType, String strValue)
	{
		switch(strType)
		{
			case "Input Text":
				driver.findElement(By.id(strElement)).sendKeys(strValue);;				
				break;
				
			case "Button":
				driver.findElement(By.id(strElement)).click();
				break;
			
			case "Frame":
				driver.findElement(By.id(strElement)).isSelected();
				break;
		}
	}
	
	public void iWebDriverElementXpath(String strElement, String strType, String strValue)
	{
		switch(strType)
		{
			case "Input Text":
				driver.findElement(By.xpath(strElement)).sendKeys((strValue));
				break;
				
			case "Button":
				driver.findElement(By.xpath(strElement)).click();
				break;
				
			case "Frame":
				driver.findElement(By.xpath(strElement)).isSelected();
				break;
		}
	}
	
	
	//************************************************************************************************************************************
	//Architecture Test Cases
	//************************************************************************************************************************************
	public static void setArchTestCaseId(String strArchMaptestcaseid)
	{
		ArchMaptestcaseid = strArchMaptestcaseid;
	}
	
	public static String getArchTestCaseId()
	{
		return ArchMaptestcaseid;
	}
	
	public static void setArchTestCaseDescription(String strArchMaptestcasedesciption)
	{
		ArchMaptestcasedesciption = strArchMaptestcasedesciption;
	}
	
	public static String getArchTestCaseDescription()
	{
		return ArchMaptestcasedesciption;
	}
	
	public static void setArchFuncCode(String strArchMapfunccode)
	{
		ArchMapfunccode = strArchMapfunccode;
	}
	
	public static String getArchFuncCode()
	{
		return ArchMapfunccode;
	}
	
	public static void setArchIteration(int intArchMapiteration)
	{
		ArchMapiteration = intArchMapiteration;
	}
	
	public static int getArchIteration()
	{
		return ArchMapiteration;
	}
	
	public static void setArchExecuteMapFlag(String strArchMapexecutionflag)
	{
		ArchMapexecutionflag = strArchMapexecutionflag;
	}
	
	public static String getArchExecuteMapFlag()
	{
		return ArchMapexecutionflag;
	}
	
	public static void setArchResult(String strArchMapresult)
	{
		ArchMapresult = strArchMapresult;
	}
	
	public static String getArchResult()
	{
		return ArchMapresult;
	}
	
	public static void setArchTestCycle(int intArchTestCycle)
	{
		ArchMaptestcycle = intArchTestCycle;
	}
	
	public static int getArchTestCycle()
	{
		return ArchMaptestcycle;
	}
	
	public static void setArchTestRun(int intArchTestRun)
	{
		ArchMaptestrun = intArchTestRun;
	}
	
	public static int getArchTestRun()
	{
		return ArchMaptestrun;
	}
	
	
	
	//************************************************************************************************************************************

	//************************************************************************************************************************************
	public static void ArchitectureTestMapping(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			//Test Case ID
			ArchMaptestcaseid = givenValue;
			break;
			
		case 1:
			//Test Case Description
			ArchMaptestcasedesciption = givenValue;
			break;
			
		case 2:
			//Iteration
			ArchMapiteration = Integer.parseInt(givenValue);
			break;
			
		case 3:
			//Execution Flag
			ArchMapexecutionflag = givenValue;
			break;
			
		case 4:
			//Result
			ArchMapresult = givenValue;
			break;
			
		case 5:
			ArchMaptestcycle = Integer.parseInt(givenValue);
			break;
			
		case 6:
			ArchMaptestrun = Integer.parseInt(givenValue);
			break;
			
			
		}
	}
	//************************************************************************************************************************************
	
	//************************************************************************************************************************************
	//Architecture Automation App Data Source
	//************************************************************************************************************************************
	public static void setDataSourceTestCaseId(String strDataSourceTestCaseId)
	{
		DataSourcetestcaseid = strDataSourceTestCaseId;
	}
	
	public static String getDataSourceTestCaseId()
	{
		return DataSourcetestcaseid;
	}
	
	public static void setDataSourceUser(String strDataSourceUser)
	{
		DataSourceuser = strDataSourceUser;
	}
	
	public static String getDataSourceuser()
	{
		return DataSourceuser;
	}
	
	public static void setDataSourceMainModule(String strDataSourceMainModule)
	{
		DataSourcemainmodule = strDataSourceMainModule;
	}
	
	public static String getDataSourceMainModule()
	{
		return DataSourcemainmodule;
	}
	
	public static void setDataSourceSubModule(String strDataSourceSubModule)
	{
		DataSourcesubmodule = strDataSourceSubModule;
	}
	
	public static String getDataSourceSubModule()
	{
		return DataSourcesubmodule;
	}
	
	public static void setDataSourceFunctionCode(String strDataSourceFunctionCode)
	{
		DataSourcefunctioncode = strDataSourceFunctionCode;
	}
	
	public static String getDataSourceFunctionCode()
	{
		return DataSourcefunctioncode;
	}
	
	public static void setDataSourceFunctionality(String strDataSourceFunctionality)
	{
		DataSourcefunctionality = strDataSourceFunctionality;
	}
	
	public static String getDataSourceFunctionality()
	{
		return DataSourcefunctionality;
	}
	
	public static void setDataSourceMainMenu(String strDataSourceMainMenu)
	{
		DataSourcemainmenu = strDataSourceMainMenu;
	}
	
	public static String getDataSourceMainMenu()
	{
		return DataSourcemainmenu;
	}
	
	public static void setDataSourceSubMenu(String strDataSourceSubMenu)
	{
		DataSourcesubmenu = strDataSourceSubMenu;
	}
		
	public static String getDataSourceSubMenu()
	{
		return DataSourcesubmenu;
	}
	//************************************************************************************************************************************
	
	//************************************************************************************************************************************
	public static void ArchitecttureTestDataSource(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			//Test Case ID
			DataSourcetestcaseid = givenValue;
			break;
			
		case 1:
			//User
			DataSourceuser = givenValue;
			break;
			
		case 2:
			//Main Modules
			DataSourcemainmodule = givenValue;
			break;
			
		case 3:
			//Sub Modules
			DataSourcesubmodule = givenValue;
			break;
			
		case 4:
			//Function Code
			DataSourcefunctioncode = givenValue;
			break;
			
		case 5:
			//Functionality
			DataSourcefunctionality = givenValue;
			break;
			
		case 6:
			//Main Menu
			DataSourcemainmenu = givenValue;
			break;
			
		case 7:
			//Sub Menu
			DataSourcesubmenu = givenValue;
			break;
		}
	}
	
	//************************************************************************************************************************************
	//Function Objects
	//************************************************************************************************************************************
	public static void setFuncModule(String strFuncmodule)
	{
		Funcmodule = strFuncmodule;
	}
	
	public static String getFuncModule()
	{
		return Funcmodule;
	}
	
	public static void setFuncField(String strFuncfield)
	{
		Funcfield = strFuncfield;
	}
	
	public static String getFuncfield()
	{
		return Funcfield;
	}
	
	public static void setFuncElementid(String strFuncelementid)
	{
		Funcelementid = strFuncelementid;
	}
	
	public static String getFuncElementid()
	{
		return Funcelementid;
	}
	
	public static void setFuncElementxpath(String strFuncelementxpath)
	{
		Funcelementxpath = strFuncelementxpath;
	}
	
	public static String getFuncElementxpath()
	{
		return Funcelementxpath;
	}
	
	public static void setFuncLinktext(String strFunclinktext)
	{
		Funclinktext = strFunclinktext;
	}
	
	public static String getFuncLinktext()
	{
		return Funclinktext;
	}
	
	public static void setFuncPageurl(String strFuncpageurl)
	{
		Funcpageurl = strFuncpageurl;
	}
	
	public static String getFuncPageurl()
	{
		return Funcpageurl;
	}
	
	public static void setFuncTitle(String strFunctitle)
	{
		Functitle = strFunctitle;
	}
	
	public static String getFunctitle()
	{
		return Functitle;
	}
	
	public static void setFuncElementtype(String strFuncelementtype)
	{
		Funcelementtype = strFuncelementtype;
	}
	
	public static String getFuncElementtype()
	{
		return Funcelementtype;
	}
	
	public static void setFuncValue(String strFuncvalue)
	{
		Funcvalue = strFuncvalue;
	}
	
	public static String getFuncValue()
	{
		return Funcvalue;
	}
	
	public static void setFuncAction(String strFuncaction)
	{
		Funcaction = strFuncaction;
	}
	
	public static String getFuncAction()
	{
		return Funcaction;
	}
	
	public static void setFuncMenureference(String strFuncmenureference)
	{
		Funcmenureference = strFuncmenureference;
	}
	
	public static String getFuncMenureference()
	{
		return Funcmenureference;
	}
	
	public static void setFuncWorksheetreference(String strFuncworksheetreference)
	{
		Funcworksheetreference = strFuncworksheetreference;
	}
	
	public static String getFuncWorksheetreference()
	{
		return Funcworksheetreference;
	}
	
	public static void setFuncFieldreference(String strFuncfieldreference)
	{
		Funcfieldreference = strFuncfieldreference;
	}
	
	public static String getFuncFieldreference()
	{
		return Funcfieldreference;
	}
	
	public static void setFuncMessage(String strFuncmessage)
	{
		Funcmessage = strFuncmessage;
	}
	
	public static String getFuncMessage()
	{
		return Funcmessage;
	}
	
	public static void setFuncResult(String strFuncresult)
	{
		Funcresult = strFuncresult;
	}
	
	public static String getFuncResult()
	{
		return Funcresult;
	}
	//************************************************************************************************************************************

	//************************************************************************************************************************************
	public static void TDDFunction(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			//Module
			Funcmodule = givenValue;
			break;
			
		case 1:
			//Field
			Funcfield = givenValue;
			break;
			
		case 2:
			//Element ID
			Funcelementid = givenValue;
			break;
			
		case 3:
			//Element Xpath
			Funcelementxpath = givenValue;
			break;
			
		case 4:
			//Link Text
			Funclinktext = givenValue;
			break;
			
		case 5:
			//Page Url
			Funcpageurl = givenValue;
			break;
			
		case 6:
			//Title
			Functitle = givenValue;
			break;
			
		case 7:
			//Element Type
			Funcelementtype = givenValue;
			break;
			
		case 8:
			//Value
			Funcvalue = givenValue;
			break;
			
		case 9:
			//Action
			Funcaction = givenValue;
			break;
			
		case 10:
			//Menu Reference
			Funcmenureference = givenValue;
			break;
			
		case 11:
			//Worksheet Reference
			Funcworksheetreference = givenValue;
			break;
			
		case 12:
			//Field Reference
			Funcfieldreference = givenValue;
			break;
			
		case 13:
			//Message
			Funcmessage = givenValue;
			break;
			
		case 14:
			//Result
			Funcresult = givenValue;
			break;
			
		}
	}

	//************************************************************************************************************************************
	
	//************************************************************************************************************************************
	// Universal Teller Menu
	//************************************************************************************************************************************
	public static void setMenumodule(String givenValue)
	{
		Menumodule = givenValue;
	}
	
	public static String getMenumodule()
	{
		return Menumodule;
	}
	
	public static void setMenuWorksheetCodeMain(String givenValue)
	{
		Menuworksheetcode_main = givenValue;
	}
	
	public static String getMenuWorksheetCodeMain()
	{
		return Menuworksheetcode_main;
	}
	
	public static void setMenuWorksheetCodeSub(String givenValue)
	{
		Menuworksheetcode_sub = givenValue;
	}
	
	public static String getMenuWorksheetCodeSub()
	{
		return Menuworksheetcode_sub;
	}
	
	public static void setMenuWorksheet(String givenValue)
	{
		Menuworksheet = givenValue;
	}
	
	public static String getMenuWorksheet()
	{
		return Menuworksheet;
	}
	
	//************************************************************************************************************************************
	public static void UnivTellerMenu(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			Menumodule = givenValue;
			break;
			
		case 1:
			Menuworksheetcode_main = givenValue;
			break;
			
		case 2:
			Menuworksheetcode_sub = givenValue;
			break;
			
		case 3:
			Menuworksheet = givenValue;
			break;
		}
	}
	//************************************************************************************************************************************
	
	//************************************************************************************************************************************
	// Universal Teller Menu - Main menu
	//************************************************************************************************************************************
	public static void setMainMenuModule(String givenValue)
	{
		MainMenumodule = givenValue;
	}
	
	public static String getMainMenuModule()
	{
		return SubMenumodule;
	}
	
	public static void setMainMenuField(String givenValue)
	{
		MainMenufield = givenValue;
	}
	
	public static String getMainMenuField()
	{
		return MainMenufield;
	}
	
	public static void setMainMenuElementId(String givenValue)
	{
		MainMenuelementid = givenValue;
	}
	
	public static String getMainMenuElementId()
	{
		return MainMenuelementid;
	}
	
	public static void setMainMenuElementXpath(String givenValue)
	{
		MainMenuelementxpath = givenValue;
	}
	
	public static String getMainMenuElementXpath()
	{
		return MainMenuelementxpath;
	}
	
	public static void setMainMenuLinkText(String givenValue)
	{
		MainMenulinktext = givenValue;
	}
	
	public static String getMainMenuLinkText()
	{
		return MainMenulinktext;
	}
	
	public static void setMainMenuPageUrl(String givenValue)
	{
		MainMenupageurl = givenValue;
	}
	
	public static String getMainMenuPageUrl()
	{
		return MainMenupageurl;
	}
	
	public static void setMainMenuTitle(String givenValue)
	{
		MainMenutitle = givenValue;
	}
	
	public static String getMainMenuTitle()
	{
		return MainMenutitle;
	}
	
	public static void setMainMenuElementType(String givenValue)
	{
		MainMenuelementtype = givenValue;
	}
	
	public static String getMainMenuElementType()
	{
		return MainMenuelementtype;
	}
	
	public static void setMainMenuValue(String givenValue)
	{
		MainMenuvalue = givenValue;
	}
	
	public static String getMainMenuValue()
	{
		return MainMenuvalue;
	}
	
	public static void setMainMenuAction(String givenValue)
	{
		MainMenuaction = givenValue;
	}
	
	public static String getMainMenuAction()
	{
		return MainMenuaction;
	}
	
	public static void setMainMenuMenuReference(String givenValue)
	{
		MainMenumenureference = givenValue;
	}
	
	public static String getMainMenuMenuReference()
	{
		return MainMenumenureference;
	}
	
	public static void setMainMenuWorksheetReference(String givenValue)
	{
		MainMenuworksheetreference = givenValue;
	}
	
	public static String getMainMenuWorksheetReference()
	{
		return MainMenuworksheetreference;
	}
	
	public static void setMainMenuFieldReference(String givenValue)
	{
		MainMenufieldreference = givenValue;
	}
	
	public static String getMainMenuFieldReference()
	{
		return MainMenufieldreference;
	}
	
	public static void setMainMenuMessage(String givenValue)
	{
		MainMenumessage = givenValue;
	}
	
	public static String getMainMenuMessage()
	{
		return MainMenumessage;
	}
	
	public static void setMainMenuResult(String givenValue)
	{
		MainMenuresult = givenValue;
	}
	
	public static String getMainMenuResult()
	{
		return MainMenuresult;
	}
	
	public void UnivTellerMenu_Main(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			MainMenumodule = givenValue;
			break;
			
		case 1:
			MainMenufield = givenValue;
			break;
			
		case 2:
			MainMenuelementid = givenValue;
			break;
			
		case 3:
			MainMenuelementxpath = givenValue;
			break;
			
		case 4:
			MainMenulinktext = givenValue;
			break;
			
		case 5:
			MainMenupageurl = givenValue;
			break;
			
		case 6:
			MainMenutitle = givenValue;
			break;
			
		case 7:
			MainMenuelementtype = givenValue;
			break;
			
		case 8:
			MainMenuvalue = givenValue;
			break;
			
		case 9:
			MainMenuaction = givenValue;
			break;
			
		case 10:
			MainMenumenureference = givenValue;
			break;
			
		case 11:
			MainMenuworksheetreference = givenValue;
			break;
			
		case 12:
			MainMenufieldreference = givenValue;
			break;
			
		case 13:
			MainMenumessage = givenValue;
			break;
			
		case 14:
			MainMenuresult = givenValue;
			break;
		}
	}
	
	//************************************************************************************************************************************
	
	//************************************************************************************************************************************
	// Universal Teller Menu - Sub menu
	//************************************************************************************************************************************
	public static void setSubMenuModule(String strSubMenuModule)
	{
		SubMenumodule = strSubMenuModule;
	}
	
	public static String getSubMenuModule()
	{
		return SubMenumodule;
	}
	
	public static void setSubMenuField(String strSubMenuField)
	{
		SubMenufield = strSubMenuField;
	}
	
	public static String getSubMenuField()
	{
		return SubMenufield;
	}
	
	public static void setSubMenuElementId(String strSubMenuElementId)
	{
		SubMenuelementid = strSubMenuElementId;
	}
	
	public static String getSubMenuElementId()
	{
		return SubMenuelementid;
	}
	
	public static void setSubMenuElementXpath(String strSubMenuElementXpath)
	{
		SubMenuelementxpath = strSubMenuElementXpath;
	}
	
	public static String getSubMenuElementXpath()
	{
		return SubMenuelementxpath;
	}
	
	public static void setSubMenuLinkText(String strSubMenuLinkText)
	{
		SubMenulinktext = strSubMenuLinkText;
	}
	
	public static String getSubMenuLinkText()
	{
		return SubMenulinktext;
	}
	
	public static void setSubMenuPageUrl(String strSubMenuPageUrl)
	{
		SubMenupageurl = strSubMenuPageUrl;
	}
	
	public static String getSubMenuPageUrl()
	{
		return SubMenupageurl;
	}
	
	public static void setSubMenuTitle(String strSubMenuTitle)
	{
		SubMenutitle = strSubMenuTitle;
	}
	
	public static String getSubMenuTitle()
	{
		return SubMenutitle;
	}
	
	public static void setSubMenuElementType(String strSubMenuElementType)
	{
		SubMenuelementtype = strSubMenuElementType;
	}
	
	public static String getSubMenuElementType()
	{
		return SubMenuelementtype;
	}
	
	public static void setSubMenuValue(String strSubMenuValue)
	{
		SubMenuvalue = strSubMenuValue;
	}
	
	public static String getSubMenuValue()
	{
		return SubMenuvalue;
	}
	
	public static void setSubMenuAction(String strSubMenuAction)
	{
		SubMenuaction = strSubMenuAction;
	}
	
	public static String getSubMenuAction()
	{
		return SubMenuaction;
	}
	
	public static void setSubMenuMenuReference(String strSubMenuMenuReference)
	{
		SubMenumenureference = strSubMenuMenuReference;
	}
	
	public static String getSubMenuMenuReference()
	{
		return SubMenumenureference;
	}
	
	public static void setSubMenuWorksheetReference(String strSubMenuWorksheetReference)
	{
		SubMenuworksheetreference = strSubMenuWorksheetReference;
	}
	
	public static String getSubMenuWorksheetReference()
	{
		return SubMenuworksheetreference;
	}
	
	public static void setSubMenuFieldReference(String strSubMenuFieldReference)
	{
		SubMenufieldreference = strSubMenuFieldReference;
	}
	
	public static String getSubMenuFieldReference()
	{
		return SubMenufieldreference;
	}
	
	public static void setSubMenuMessage(String strSubMenuMessage)
	{
		SubMenumessage = strSubMenuMessage;
	}
	
	public static String getSubMenuMessage()
	{
		return SubMenumessage;
	}
	
	public static void setSubMenuResult(String strSubMenuResult)
	{
		SubMenuresult = strSubMenuResult;
	}
	
	public static String getSubMenuResult()
	{
		return SubMenuresult;
	}
	
	public void UnivTellerMenu_Sub(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			SubMenumodule = givenValue;
			break;
			
		case 1:
			SubMenufield = givenValue;
			break;
			
		case 2:
			SubMenuelementid = givenValue;
			break;
			
		case 3:
			SubMenuelementxpath = givenValue;
			break;
			
		case 4:
			SubMenulinktext = givenValue;
			break;
			
		case 5:
			SubMenupageurl = givenValue;
			break;
			
		case 6:
			SubMenutitle = givenValue;
			break;
			
		case 7:
			SubMenuelementtype = givenValue;
			break;
			
		case 8:
			SubMenuvalue = givenValue;
			break;
			
		case 9:
			SubMenuaction = givenValue;
			break;
			
		case 10:
			SubMenumenureference = givenValue;
			break;
			
		case 11:
			SubMenuworksheetreference = givenValue;
			break;
			
		case 12:
			SubMenufieldreference = givenValue;
			break;
			
		case 13:
			SubMenumessage = givenValue;
			break;
			
		case 14:
			SubMenuresult = givenValue;
			break;
		}
	}
	//************************************************************************************************************************************
	
	
	//************************************************************************************************************************************
	public static void setScriptModlue(String strScriptModule)
	{
		Scriptmodule = strScriptModule;
	}
	
	public static String getScriptModule()
	{
		return Scriptmodule;
	}
	
	public static void setScriptField(String strScriptField)
	{
		Scriptfield = strScriptField;
	}
	
	public static String getScriptField()
	{
		return Scriptfield;
	}
	
	public static void setScriptElementId(String strScriptElementId)
	{
		Scriptelementid = strScriptElementId;
	}
	
	public static String getScriptElementId()
	{
		return Scriptelementid;
	}
	
	public static void setScriptElementXpath(String strScriptElementXpath)
	{
		Scriptelementxpath = strScriptElementXpath;
	}
	
	public static String getScriptElementXpath()
	{
		return Scriptelementxpath;
	}
	
	public static void setScriptLinkText(String strScriptLinkText)
	{
		Scriptlinktext = strScriptLinkText;
	}
	
	public static String getScriptLinkText()
	{
		return Scriptlinktext;
	}
	
	public static void setScriptPageUrl(String strScriptPageUrl)
	{
		Scriptpageurl = strScriptPageUrl;
	}
	
	public static String getScriptPageUrl()
	{
		return Scriptpageurl;
	}
	
	public static void setScriptTitle(String strScriptTitle)
	{
		Scripttitle = strScriptTitle;
	}
	
	public static String getScriptTitle()
	{
		return Scripttitle;
	}
	
	public static void setScriptElementType (String strScriptElementType)
	{
		Scriptelementtype = strScriptElementType;
	}
	
	public static String getScriptElementType()
	{
		return Scriptelementtype;
	}
	
	public static void setScriptValue(String strScriptValue)
	{
		 Scriptvalue = strScriptValue;
	}
	
	public static String getScriptValue()
	{
		return Scriptvalue;
	}
	
	public static void setScriptAction(String strScriptAction)
	{
		Scriptaction = strScriptAction;
	}
	
	public static String getScriptAction()
	{
		return Scriptaction;
	}
	
	public static void setScriptMenuReference(String strScriptMenuReference)
	{
		Scriptmenureference = strScriptMenuReference;
	}
	
	public static String getScriptMenuReference()
	{
		return Scriptmenureference;
	}
	
	public static void setScriptWorksheetReference(String strScriptWorksheetReference)
	{
		Scriptworksheetreference = strScriptWorksheetReference;
	}
	
	public static String getScriptWorksheetReference()
	{
		return Scriptworksheetreference;
	}
	
	public static void setScriptFieldReference(String strScriptFieldReference)
	{
		Scriptfieldreference = strScriptFieldReference;
	}
	
	public static String getScriptFieldReference()
	{
		return Scriptfieldreference;
	}
	
	public static void setScriptMessage(String strScriptMessage)
	{
		Scriptmessage = strScriptMessage;
	}
	
	public static String getScriptMessage()
	{
		return Scriptmessage;
	}
	
	public static void setScriptResult(String strScriptResult)
	{
		Scriptresult = strScriptResult;
	}
	
	public static String getScriptResult()
	{
		return Scriptresult;
	}
	
	public static void MapScript(int flag, String givenValue)
	{
		switch(flag)
		{
		case 0:
			Scriptmodule = givenValue;
			break;
			
		case 1:
			Scriptfield = givenValue;
			break;
			
		case 2:
			Scriptelementid = givenValue;
			break;
			
		case 3:
			Scriptelementxpath = givenValue;
			break;
			
		case 4:
			Scriptlinktext = givenValue;
			break;
			
		case 5:
			Scriptpageurl = givenValue;
			break;
			
		case 6:
			Scripttitle = givenValue;
			break;
			
		case 7:
			Scriptelementtype = givenValue;
			break;
			
		case 8:
			Scriptvalue = givenValue;
			break;
			
		case 9:
			Scriptaction = givenValue;
			break;
			
		case 10:
			Scriptmenureference = givenValue;
			break;
			
		case 11:
			Scriptworksheetreference = givenValue;
			break;
			
		case 12:
			Scriptfieldreference = givenValue;
			break;
			
		case 13:
			Scriptmessage = givenValue;
			break;
			
		case 14:
			Scriptresult = givenValue;
			break;
		}
	}
	
	
	//************************************************************************************************************************************
	
	//************************************************************************************************************************************
	public void arrayTestResultMap(int arrayMapID, String arrayValue)
	{
		
	}
}
