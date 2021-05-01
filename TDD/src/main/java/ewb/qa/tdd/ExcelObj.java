/*
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;
import java.util.*;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;

import org.openqa.selenium.*;

import java.time.LocalDateTime;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
*/
package ewb.qa.tdd;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;
import java.util.*;

import org.apache.commons.io.FileUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;

import org.openqa.selenium.*;

import java.time.LocalDateTime;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class ExcelObj {

    public static XSSFWorkbook wbMain;
    public static XSSFWorkbook wbReference;
    public static XSSFSheet shMain;
    public static XSSFSheet shReference;
    public static FileInputStream fis;
    public static FileOutputStream fos;
    public static XSSFRow rowMain;
    public static Cell cellMain;
    public static XSSFRow rowReference;
    public static Cell cellReference;
    public static File file;

    //D:\\QA\\Projects\\Test Automation\\VS_SeleniumWebDriver\\QA_SeleniumWebDriver\\TDD.xlsx
    public String strPath = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TDD.xlsx";
    public int lastNumber;
    public String sheetName;

    private String strbrowser;
    private String strbaseurl;
    private String strmodule;
    private String strfield;
    private String strelementid;
    private String strelementxpath;
    private String strlinktext;
    private String strpageurl;
    private String strtitle;
    private String strelementtype;
    private String strvalue;
    private String straction;
    private static String strreferencesheet;
    private String strreferencefield;
    private static String strmessage;
    private static String strcustomernumber;
    public String strsheetname;

    private String strtestcaseid;
    private String strtestcasedesc;
    private int intiteration;
    private String strexecuteflag;
    private String strtestcaseresult;
    private String strworksheet;

    static SeleniumObj seleniumObj;
    public Properties propertiesObj;


    /**
     * @param Module
     * @param Field
     * @param Elementtype
     * @param Elementid
     * @param Elementxpath
     * @param Elementvalue
     */
    /*
    public static boolean MapElementType(String Module, String Field, String Title, String Elementtype, String Elementid, String Elementxpath, String Elementvalue, String Action, String ReferenceField, boolean TakeScreenshot, String testcaseid){
            seleniumObj = new SeleniumObj();
            boolean MapElementResult = true;
            try
            {
                    switch(Elementtype)
                    {
                    case "Input Box":

                            if(!Elementvalue.equals("null") && !ReferenceField.equals("null"))
                            {
                                    StoreToTemp(Elementvalue, ReferenceField);
                            }
                            else if(Elementvalue.equals("null") && !ReferenceField.equals("null"))
                            {
                                    Elementvalue = ReadToTemp(ReferenceField);
                            }

                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.SendKeys_EventById(Elementid, Elementvalue, ReferenceField);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    seleniumObj.SendKeys_EventByXpath(Elementxpath, Elementvalue, ReferenceField);
                            }

                            break;

                    case "Button":
                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.Click_EventById(Elementid);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    seleniumObj.Click_EventByXpath(Elementxpath);
                            }
                            break;

                    case "Text":
                            String browserMessage = "";
                            if(!Elementid.equals("null"))
                            {
                                    browserMessage = seleniumObj.Text_EventById(Elementid);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    browserMessage = seleniumObj.Text_EventByXpath(Elementxpath);
                            }

                            if(browserMessage.equalsIgnoreCase(strmessage))
                            {
                                    break;
                            }
                            else
                            {
                                    break;
                            }

                    case "Frame":
                            int iframeNum = Integer.parseInt(Elementvalue);

                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.Select_EventById(Elementid, iframeNum);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    seleniumObj.Select_EventByXpath(Elementxpath, iframeNum);
                            }
                            break;

                    case "Browse":
                            String strTitleNewTab;
                            strTitleNewTab = seleniumObj.Title_NewTab();
                            System.out.println(strTitleNewTab);
                            break;

                    case "NewTab":
                            switch(Action)
                            {
                            case "SwitchTo":
                                    seleniumObj.NewTab_Switchto();
                                    break;

                            case "Close":
                                    seleniumObj.NewTab_Close();
                                    break;
                            }
                            break;

                    case "SubTab":
                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.SubTab_SelectById(Elementid);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    seleniumObj.SubTab_SelectByXpath(Elementxpath);
                            }
                            break;

                    case "Submit":
                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.Submit_EventById(Elementid);				
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    seleniumObj.Submit_EventByXpath(Elementxpath);
                            }
                            break;

                    case "FrameOut":
                            seleniumObj.Switchto_Default();
                            break;

                    case "Dropdown Box":
                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.DropdownBox_ById(Elementid, Elementvalue, ReferenceField);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    seleniumObj.DropdownBox_ByXpath(Elementxpath, Elementvalue, ReferenceField);
                            }
                            break;

                    case "Text Message":
                            if(!Elementid.equals("null"))
                            {
                                    switch(Field)
                                    {
                                    case "Customer Number":
                                            strcustomernumber = seleniumObj.Text_EventById(Elementid);

                                            System.out.println("Customer Number : " + strcustomernumber);
                                            StoreToTemp(strcustomernumber, ReferenceField);
                                            break;

                                    case "Transaction Complete":
                                            String createdCustomerNumber;
                                            createdCustomerNumber = seleniumObj.Text_EventById(Elementid);
                                            System.out.println("Transaction Completed");
                                            break;
                                    }

                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    switch(Field)
                                    {
                                    case "Customer Number":
                                            strcustomernumber = seleniumObj.Text_EventByXpath(Elementxpath);

                                            System.out.println("Customer Number : " + strcustomernumber);
                                            StoreToTemp(strcustomernumber, ReferenceField);
                                            break;

                                    case "Transaction Complete":
                                            String createdCustomerNumber;
                                            createdCustomerNumber = seleniumObj.Text_EventByXpath(Elementxpath);
                                            System.out.println("Transaction Completed");
                                            break;
                                    }

                            }
                            break;
                            
                    case "Radio":
                            if(!Elementid.equals("null"))
                            {
                                    seleniumObj.Click_EventById(Elementid);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                    if(Field.contains("GENDER")){
                                            Elementxpath = "//*[@value=" + '"' + Elementvalue + '"' + "]";
                                    }                                
                                    seleniumObj.Click_EventByXpath(Elementxpath);
                            }
                            break;                        
                    }			

                    if(TakeScreenshot == true)
                    {
                            TakeScreenShot(seleniumObj.driver, testcaseid, Module, Field);
                    }
            }
            catch(Exception ex)
            {
                    System.out.println(ex.getMessage().toString());
                    MapElementResult = false;
            }
            return MapElementResult;
    }
    */
    
    public static void getArchTestCases(String filename, String worksheetname) throws IOException{
            Properties propObj = new Properties();
            SQLObj sqlObj = new SQLObj();

            FileInputStream fis = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet worksheet = workbook.getSheet(worksheetname);
            int iExcelLastRow = worksheet.getLastRowNum();

            String browser = worksheet.getRow(0).getCell(1).getStringCellValue();
            String baseurl = worksheet.getRow(1).getCell(1).getStringCellValue();

            propObj.setBrowser(browser);
            propObj.setBaseUrl(baseurl);
            String commandValue = "";

            try
            {
                Connection conn = null;
                conn = sqlObj.ConnToDB();

                for(int iExcelRow=3; iExcelRow<=iExcelLastRow; iExcelRow++)
                {
                    CallableStatement cStmt = conn.prepareCall("{call Insert_TestCases(?,?,?,?,?,?,?)}");

                    String conCatValue = "";
                    int paramIndex = 1;

                    for(int iExcelCell=0; iExcelCell<=6; iExcelCell++)
                    {
                        XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                        String cellStringValue = "";
                        int cellIntValue = 0;

                        if(cellObj == null)
                        {
                                cellStringValue = null;
                        }
                        else
                        {
                            switch(cellObj.getCellType())
                            {
                                case STRING:
                                        //cellValue = "STRING" + "##" + cellObj.getStringCellValue();
                                        cellStringValue = cellObj.getStringCellValue();
                                        cStmt.setString(paramIndex, cellStringValue);
                                        break;

                                case NUMERIC:
                                        //cellValue = "INT" + "##" + NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                        cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                        cStmt.setInt(paramIndex, cellIntValue);
                                        break;

                                case BLANK:
                                        //cellValue = "STRING" + "##" + null;
                                        cellStringValue = "null";
                                        cStmt.setString(paramIndex, cellStringValue);
                                        break;

                                case ERROR:
                                        //cellValue = "STRING" + "##" + null;
                                        cellStringValue = "null";
                                        cStmt.setString(paramIndex, cellStringValue);
                                        break;
                            }
                        }
                        paramIndex++;
                    }
                    cStmt.execute();
                    cStmt.close();
                }
                fis.close();
                conn.close();
            }
            catch(Exception ex)
            {
                                System.out.println("Error Message: " + ex.getMessage() + "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + "\n" + "Stack Trace: " + ex.getStackTrace() + "\n" + "Cause: " + ex.getCause());
            }

    }

    //*********************************************************************************************************************
    /*
    public static void readArchTestCases() throws IOException
    {
            Properties propObj = new Properties();
            SQLObj sqlObj = new SQLObj();
            Connection conn = null;
            CallableStatement cStmt = null;

            try
            {
                    conn = sqlObj.ConnToDB();
                    cStmt = conn.prepareCall("{call Search_TestCaseID_InTestCase()}");
                    ResultSet rs = cStmt.executeQuery();

                    while(rs.next())
                    {
                            String testcaseid = rs.getNString("TESTCASEID");
                            String executeflag = rs.getNString("EXECUTEFLAG");
                            System.out.println(testcaseid);

                            String filename = "";
                            String worksheetname = "";

                            if(executeflag.equalsIgnoreCase("Y"))
                            {
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\T24 Architecture Layout.xlsx";
                                    worksheetname = "Automation App DataSource";
                                    getArchDataSource(filename, worksheetname);

                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TDD_FunctionalityMaker.xlsx";
                                    getTDDFunc(filename);
                                    //readArchDataSource(testcaseid);

                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TDD_FunctionalityAuthoriser.xlsx";
                                    getTDDFunc(filename);

                            }
                    }

                    conn.close();	
            }
            catch(Exception ex)
            {
                    System.out.println("Error Message: " + ex.getMessage() + "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + "\n" + "Stack Trace: " + ex.getStackTrace() + "\n" + "Cause: " + ex.getCause());
            }
    }
    */
    //*********************************************************************************************************************

    public static void getArchDataSource(String filename, String worksheetname, String spcommand, String qcommand, int lastcolumn) throws IOException
    {
            Properties propObj = new Properties();
            //ArrayList<String> arrArchAutoAppDataSource = new ArrayList();
            FileInputStream fis = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet worksheet = workbook.getSheet(worksheetname);
            int iExcelLastRow = worksheet.getLastRowNum() - 1;
            int iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;

            SQLObj sqlObj = new SQLObj();
            Connection conn = null;
            CallableStatement cStmt = null;
            Statement Stmt = null;

            try
            {
                    conn = sqlObj.ConnToDB();
                    Stmt = conn.createStatement();
                    //int intStmt = Stmt.executeUpdate("TRUNCATE TABLE dbo.APP_DATASOURCE");
                    int intStmt = Stmt.executeUpdate(qcommand);
                    conn.commit();

                    if(intStmt == -1)
                    {
                            for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                            {
                                    //cStmt = conn.prepareCall("{call Insert_AppDataSource(?,?,?,?,?,?,?,?,?,?)}");
                                    cStmt = conn.prepareCall(spcommand);
                                    String conCatValue = "";
                                    int paramindex = 1;
                                    int iRow = 0;

                                    for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                    {
                                            XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                                            String cellStringValue = "";
                                            int cellIntValue = 0;

                                            if(cellObj == null)
                                            {
                                                    cellStringValue = cellObj.getStringCellValue();
                                                    cStmt.setString(paramindex, cellStringValue);
                                            }
                                            else
                                            {	
                                                    switch(cellObj.getCellType())
                                                    {
                                                            case STRING:
                                                                    //cellValue = "STRING" + "##" + cellObj.getStringCellValue();
                                                                    cellStringValue = cellObj.getStringCellValue();
                                                                    cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                    break;

                                                            case NUMERIC:
                                                                    //cellValue = "INT" + "##" + NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                                    cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                                    cStmt.setInt(paramindex, cellIntValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellStringValue + "\n";
                                                                    break;

                                                            case BLANK:
                                                                    //cellValue = "STRING" + "##" + null;
                                                                    cellStringValue = "null";
                                                                    cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                    break;

                                                            case ERROR:
                                                                    //cellValue = "STRING" + "##" + null;
                                                                    cellStringValue = "null";
                                                                    cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                    break;
                                                    }
                                            }
                                            iRow = iExcelCell;
                                            paramindex++;
                                    }

                                    if(iRow == iExcelLastCell)
                                    {
                                            cStmt.setInt(paramindex, iExcelRow);
                                            conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + iExcelRow + "\n";
                                    }

                                    System.out.println(conCatValue);
                                    cStmt.execute();
                                    cStmt.close();
                            }
                    }
                    fis.close();
                    conn.commit();
                    conn.close();
            }
            catch(Exception ex)
            {
                    System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
            }

    }


    //*********************************************************************************************************************
    /*
    public static void readArchDataSource(String strtestcasesid) throws IOException
    {
            Properties propObj = new Properties();
            SQLObj sqlObj = new SQLObj();
            Connection conn = null;

            CallableStatement cStmt = null;

            //Reading App DataSource
            try
            {
                    conn = sqlObj.ConnToDB();
                    cStmt = conn.prepareCall("{call Search_TestCaseID_InAppDataSource(?)}");
                    cStmt.setString(1, strtestcasesid);
                    ResultSet rs = cStmt.executeQuery();

                    while(rs.next())
                    {
                            String id = rs.getNString("ID");
                            String testcaseid = rs.getNString("TESTCASEID");
                            String usertype = rs.getNString("USERTYPE");
                            String mainmodule = rs.getNString("MAIN_MODULE");
                            String submodule = rs.getNString("SUB_MODULE");
                            String functioncode = rs.getNString("FUNCTION_CODE");
                            String functionmap = rs.getNString("FUNCTION_MAP");
                            String mainmenu = rs.getNString("MAIN_MENU");
                            String submenu = rs.getNString("SUB_MENU");
                            String childmenu = rs.getNString("CHILD_MENU");


                            String filename = "D:\\\\QA\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TestDataTemplate.xlsx";
                            String worksheetname = functionmap;
                            //getTestData(filename, worksheetname);

                            if(strtestcasesid.equalsIgnoreCase(strtestcasesid))
                            {
                                    switch(usertype)
                                    {
                                    case "Maker":
                                            //Get Test Data

                                            //Get Test Script
                                            if(mainmenu.equalsIgnoreCase("null"))
                                            {
                                                    //Goto Function and get details
                                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityMaker.xlsx";
                                                    worksheetname = functioncode;
                                                    getTDDFunc(filename);
                                                    readTDDFunc(testcaseid);
                                            }
                                            else
                                            {
                                                    //Goto Universal Teller Menu and search Main and Sub Menu
                                                    //Goto Function and get details
                                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\UniversalTellerMenu.xlsx";
                                                    worksheetname = "Teller Universal Menu";
                                                    getMenu(filename, worksheetname);
                                                    readMenu(testcaseid, mainmenu, user);			

                                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityMaker.xlsx";
                                                    worksheetname = functioncode;
                                                    getTDDFunc(filename, worksheetname);
                                                    readTDDFunc(testcaseid);
                                            }
                                            break;

                                    case "Authoriser":
                                            //Get Test Data

                                            //Get Test Script
                                            if(mainmenu.equalsIgnoreCase("null"))
                                            {
                                                    //Goto Function and get details
                                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityAuthoriser.xlsx";
                                                    worksheetname = functioncode;
                                                    getTDDFunc(filename, worksheetname);
                                                    readTDDFunc(testcaseid);
                                            }
                                            else
                                            {
                                                    //Goto Universal Teller Menu and search Main and Sub Menu
                                                    //Goto Function and get details
                                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\ServiceManagerMenu.xlsx";
                                                    worksheetname = "Teller Universal Menu";
                                                    getMenu(filename, worksheetname);
                                                    readMenu(testcaseid, mainmenu, user);			

                                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityAuthoriser.xlsx";
                                                    worksheetname = functioncode;
                                                    getTDDFunc(filename, worksheetname);
                                                    readTDDFunc(testcaseid);
                                            }
                                            break;
                                    }
                            }

                    }


                    //*********************************************************************************************************************
                    for(int iArr=0; iArr<=iArrCount; iArr++)
                    {
                            cStmt = conn.prepareCall("{call Insert_AppDataSource(?,?,?,?,?,?,?,?,?,?)}");

                            String cellValue = arrDataSource.get(iArr);
                            Scanner scanner1 = new Scanner(cellValue);
                            scanner1.useDelimiter("&&");
                            int iScanner = 1;
                            while(scanner1.hasNext())
                            {
                                    String givenString = scanner1.next();

                                    Scanner scanner2 = new Scanner(givenString);
                                    scanner2.useDelimiter("##");

                                    String paramType = "";
                                    String paramValue = "";
                                    int iterate = 1;
                                    while(scanner2.hasNext())
                                    {
                                            String givenValue = scanner2.next();
                                            switch(iterate)
                                            {
                                            case 1:
                                                    paramType = givenValue;
                                                    break;

                                            case 2:
                                                    paramValue = givenValue;
                                                    break;
                                            }
                                            iterate++;
                                    }

                                    switch(paramType)
                                    {
                                    case "STRING":
                                            cStmt.setString(iScanner, paramValue);
                                            break;

                                    case "INT":
                                            cStmt.setInt(iScanner, Integer.parseInt(paramValue));
                                            break;
                                    }
                                    iScanner++;
                            }
                            cStmt.execute();
                            cStmt.close();

                            cStmt = conn.prepareCall("{call Search_TestCaseID_InAppDataSource()");
                    }
                    //*********************************************************************************************************************
            }
            catch(Exception ex)
            {
                    System.out.println("Error Message: " + ex.getMessage() + "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + "\n" + "Stack Trace: " + ex.getStackTrace() + "\n" + "Cause: " + ex.getCause());
            }


            //*********************************************************************************************************************************
            for(int iArr=0; iArr<=iArrCount; iArr++)
            {

                    String arrValue = arrDataSource.get(iArr);
                    Scanner scanner = new Scanner(arrValue);
                    scanner.useDelimiter("&&");
                    int iScan = 0;

                    while(scanner.hasNext())
                    {
                            String givenValue = scanner.next();
                            propObj.ArchitecttureTestDataSource(iScan, givenValue);
                            iScan++;
                    }

                    String testcaseid = propObj.getDataSourceTestCaseId();
                    String user = propObj.getDataSourceuser();
                    String mainmodules = propObj.getDataSourceMainModule();
                    String submodules = propObj.getDataSourceSubModule();
                    String functioncode = propObj.getDataSourceFunctionCode();
                    String functionality = propObj.getDataSourceFunctionality();
                    String mainmenu = propObj.getDataSourceMainMenu();
                    String submenu = propObj.getDataSourceSubMenu();
                    String filename = "";
                    String worksheetname = "";

                    filename = "D:\\\\QA\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TestDataTemplate.xlsx";
                    worksheetname = functioncode;
                    //getTestData(filename, worksheetname);

                    if(testcaseid.equalsIgnoreCase(strtestcasesid))
                    {
                            switch(user)
                            {
                            case "Maker":
                                    //Get Test Data

                                    //Get Test Script
                                    if(mainmenu.equalsIgnoreCase("null"))
                                    {
                                            //Goto Function and get details
                                            filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityMaker.xlsx";
                                            worksheetname = functioncode;
                                            getTDDFunc(filename, worksheetname);
                                            readTDDFunc(testcaseid);
                                    }
                                    else
                                    {
                                            //Goto Universal Teller Menu and search Main and Sub Menu
                                            //Goto Function and get details
                                            filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\UniversalTellerMenu.xlsx";
                                            worksheetname = "Teller Universal Menu";
                                            getMenu(filename, worksheetname);
                                            readMenu(testcaseid, mainmenu, user);			

                                            filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityMaker.xlsx";
                                            worksheetname = functioncode;
                                            getTDDFunc(filename, worksheetname);
                                            readTDDFunc(testcaseid);
                                    }
                                    break;

                            case "Authoriser":
                                    //Get Test Data

                                    //Get Test Script
                                    if(mainmenu.equalsIgnoreCase("null"))
                                    {
                                            //Goto Function and get details
                                            filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityAuthoriser.xlsx";
                                            worksheetname = functioncode;
                                            getTDDFunc(filename, worksheetname);
                                            readTDDFunc(testcaseid);
                                    }
                                    else
                                    {
                                            //Goto Universal Teller Menu and search Main and Sub Menu
                                            //Goto Function and get details
                                            filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\ServiceManagerMenu.xlsx";
                                            worksheetname = "Teller Universal Menu";
                                            getMenu(filename, worksheetname);
                                            readMenu(testcaseid, mainmenu, user);			

                                            filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\TDD_FunctionalityAuthoriser.xlsx";
                                            worksheetname = functioncode;
                                            getTDDFunc(filename, worksheetname);
                                            readTDDFunc(testcaseid);
                                    }
                                    break;
                            }
                    }
            }
            //*********************************************************************************************************************************

    }
    */
    //*********************************************************************************************************************


    //*********************************************************************************************************************

    public static void getGetWorksheetList_Excel(String filename, String givenworksheetname, String spcommand, int lastcolumn) throws IOException
    {
                    Properties propObj = new Properties();
                    SQLObj sqlObj = new SQLObj();

                    //ArrayList<String> arrTDDFunc = new ArrayList();

                    FileInputStream fis = new FileInputStream(filename);
                    XSSFWorkbook workbook = new XSSFWorkbook(fis);
                    XSSFSheet worksheet = null;
                    String worksheetname = "";

                    Connection conn = null;
                    CallableStatement cStmt = null;
                    Statement Stmt = null;

                    int iWorkSheets = workbook.getNumberOfSheets() - 1;
                    //int iExcelLastRow = worksheet.getLastRowNum();
                    int iExcelLastRow = 0;
                    int iExcelLastCell = 0;

                    try
                    {
                        conn = sqlObj.ConnToDB();
                        for(int iWorksheet=0; iWorksheet<=iWorkSheets; iWorksheet++)
                        {
                                String conCatValue = "";
                                worksheetname = workbook.getSheetName(iWorksheet);			
                                worksheet = workbook.getSheet(worksheetname);

                                iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;
                                if(iExcelLastCell == lastcolumn)
                                {
                                        System.out.println("Same Last Cell Count");
                                }
                                else
                                {
                                        System.out.println("Not the same " + "\n" + "iExcelLastCell: " + iExcelLastCell + "\n" + "lastcolumn: " + lastcolumn);
                                }

                                int iRow = 0;

                                //cStmt = conn.prepareCall("{call Insert_PageTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}");

                                iExcelLastRow = worksheet.getLastRowNum();

                                for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                                {
                                        int paramindex = 1;
                                        cStmt = conn.prepareCall(spcommand);
                                        cStmt.setString(paramindex, worksheetname);
                                        conCatValue = paramindex + "|" + "STRING" + "-" + worksheetname + "\n";

                                        for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                        {
                                                String cellStringValue = "";
                                                int cellIntValue = 0;
                                                paramindex++;

                                                XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                                                if(cellObj == null)
                                                {
                                                        cellStringValue = "null";
                                                        cStmt.setString(paramindex, cellStringValue);
                                                }
                                                else
                                                {
                                                        switch(cellObj.getCellType())
                                                        {
                                                        case STRING:
                                                                //cellValue = cellObj.getStringCellValue();
                                                                cellStringValue = cellObj.getStringCellValue();
                                                                cStmt.setString(paramindex, cellStringValue);
                                                                conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                break;

                                                        case NUMERIC:
                                                                //cellValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                                cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                                cStmt.setInt(paramindex, cellIntValue);
                                                                conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellIntValue + "\n";
                                                                break;

                                                        case BLANK:
                                                                //cellValue = null;
                                                                cellStringValue = "null";
                                                                cStmt.setString(paramindex, cellStringValue);
                                                                conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                break;

                                                        case ERROR:
                                                                //cellValue = null;
                                                                cellStringValue = "null";
                                                                cStmt.setString(paramindex, cellStringValue);
                                                                conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                break;
                                                        }
                                                }
                                                iRow = iExcelCell;
                                        }
                                        //arrTDDFunc.add(conCatValue);
                                        if(iRow==iExcelLastCell)
                                        {
                                                paramindex++;
                                                cStmt.setInt(paramindex, iExcelRow);
                                                conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + iExcelRow + "\n";
                                        }
                                        System.out.println(conCatValue);
                                        cStmt.execute();
                                        cStmt.close();
                                }		

                        }
                        fis.close();	
                        conn.commit();
                        conn.close();
                    }
                    catch(Exception ex)
                    {
                            System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
                    }

    }
    //*********************************************************************************************************************


    //*********************************************************************************************************************
    /*
    public static void readTDDFunc(String testcaseid) throws IOException
    {
            Properties propObj = new Properties();
            ArrayList<String> arrTDDFunc = new ArrayList();
            ArrayList<String> arrScript = new ArrayList<String>();

            arrTDDFunc = propObj.getArrTDDFunc();
            arrScript = propObj.getArrScript();
            int iScriptCount = 0;

            if(arrScript == null)
            {
                    iScriptCount = 0;
            }
            else
            {
                    iScriptCount = arrScript.size() - 1;
            }

            //System.out.println("Array List before TDD Func \n" + arrScript);

            int iArrCount = arrTDDFunc.size() - 1;

            for(int iArr = 0; iArr<=iArrCount; iArr++)
            {
                    String arrValue = arrTDDFunc.get(iArr);

                    if(arrScript == null)
                    {
                            arrScript = new ArrayList();
                            arrScript.add(iScriptCount, arrValue);				
                    }
                    else
                    {
                            iScriptCount++;
                            arrScript.add(iScriptCount, arrValue);
                    }			
            }
            propObj.setArrScript(arrScript);
    }
    */
    //*********************************************************************************************************************

    public static void getSpecificWorksheet_Excel(String filename, String worksheetname, int lastcolumn, String spcommand, String qcommand, String sequenceflag) throws IOException
    {
            Properties propObj = new Properties();
            SQLObj sqlObj = new SQLObj();
            Connection conn = null;
            CallableStatement cStmt = null;
            Statement Stmt = null;
            //int paramindex = 0;
            //ArrayList<String> arrUnivTellerMenu = new ArrayList();

            FileInputStream fis = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet worksheet = workbook.getSheet(worksheetname);
            int iExcelLastRow = worksheet.getLastRowNum();
            int iExcelLassCell = worksheet.getRow(0).getLastCellNum();

            try
            {	
                    conn = sqlObj.ConnToDB();
                    for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                    {
                            String conCatValue = "";
                            int iRow = 0;
                            int paramindex = 1;
                            //cStmt = conn.prepareCall("{call Insert_Menumap(?,?,?,?,?,?)}");
                            cStmt = conn.prepareCall(spcommand);

                            cStmt.setString(paramindex, worksheetname);
                            conCatValue = paramindex + "|" + "STRING" + "-" + worksheetname + "\n";

                            for(int iExcelCell=0; iExcelCell<=lastcolumn; iExcelCell++)
                            {	
                                    String cellStringValue = "";
                                    int cellIntValue = 0;
                                    paramindex++;

                                    XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                                    String cellValue = "";
                                    if(cellObj == null)
                                    {
                                            //cellValue = null;
                                            cellStringValue = "null";
                                            cStmt.setString(paramindex, cellStringValue);
                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                    }
                                    else
                                    {
                                            switch(cellObj.getCellType())
                                            {
                                            case STRING:
                                                    //cellValue = cellObj.getStringCellValue();
                                                    cellStringValue = cellObj.getStringCellValue();
                                                    cStmt.setString(paramindex, cellStringValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                    break;

                                            case NUMERIC:
                                                    //cellValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                    cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                    cStmt.setInt(paramindex, cellIntValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellIntValue + "\n";

                                                    break;

                                            case BLANK:
                                                    //cellValue = null;
                                                    cellStringValue = "null";
                                                    cStmt.setString(paramindex, cellStringValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                    break;

                                            case ERROR:
                                                    cellStringValue = "null";
                                                    cStmt.setString(paramindex, cellStringValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                    break;
                                            }
                                    }
                                    iRow = iExcelCell;
                            }
                            System.out.println(conCatValue);
                            cStmt.execute();
                            cStmt.close();
                    }
                    fis.close();
                    conn.commit();
                    conn.close();
            }
            catch(Exception ex)
            {
                    System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
            }
    }
    //*********************************************************************************************************************


    public static void readUserMenuMap() throws Exception
    {
            SQLObj sqlObj = new SQLObj();

            Connection conn = null;
            CallableStatement cStmt = null;
            String filename = "";
            String worksheetname = "";

            try 
            {
                    conn = sqlObj.ConnToDB();
                    cStmt = conn.prepareCall("{call Search_UserMenuMap_All()}");
                    ResultSet rs = cStmt.executeQuery();
                    int iRecordCount = rs.getRow();

                    while(rs.next())
                    {
                            String usercode = rs.getNString("USER_CODE");
                            String maincode = rs.getNString("MAIN_CODE");
                            String subcode = rs.getNString("SUB_CODE");
                            String childcode = rs.getNString("CHILD_CODE");
                            String spcommand = "{call Insert_MenuTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                            String qcommand = "";
                            String sequenceflag = "";
                            int lastcolumn = 0;

                            switch(usercode)
                            {
                            case "User_0001":
                                    //UniversalTellerMenu
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\UniversalTellerMenu.xlsx";
                                    System.out.println("User: " + usercode + "-" + "Working on worksheet: " + maincode);
                                    break;

                            case "User_0002":
                                    //ServiceManagerMenu
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\ServiceManagerMenu.xlsx";
                                    System.out.println("User: " + usercode + "-" + "Working on worksheet: " + maincode);
                                    break;
                            }

                            String conCatValue = "";

                            FileInputStream fis = new FileInputStream(filename);
                            XSSFWorkbook workbook = new XSSFWorkbook(fis);
                            worksheetname = maincode;
                            XSSFSheet worksheet = workbook.getSheet(worksheetname);

                            try
                            {
                                    int iExcelLastRow = worksheet.getLastRowNum();
                                    int iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;


                                    for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                                    {
                                            int iRow = 0;
                                            int paramindex = 2;

                                            CallableStatement ins_cStmt = null;
                                            ins_cStmt = conn.prepareCall(spcommand);

                                            ins_cStmt.setString(1, usercode);
                                            conCatValue = 1 + "|" + "STRING" + "-" + usercode + "\n";

                                            ins_cStmt.setString(2, worksheetname);
                                            conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + worksheetname + "\n";

                                            for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                            {
                                                    String cellStringValue = "";
                                                    int cellIntValue = 0;
                                                    paramindex++;

                                                    XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                                                    if(cellObj == null)
                                                    {
                                                            cellStringValue = "null";
                                                            ins_cStmt.setString(paramindex, cellStringValue);
                                                    }
                                                    else
                                                    {
                                                            switch(cellObj.getCellType())
                                                            {
                                                            case STRING:
                                                                    //cellValue = cellObj.getStringCellValue();
                                                                    cellStringValue = cellObj.getStringCellValue();
                                                                    ins_cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                    break;

                                                            case NUMERIC:
                                                                    //cellValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                                    cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                                    ins_cStmt.setInt(paramindex, cellIntValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellIntValue + "\n";
                                                                    break;

                                                            case BLANK:
                                                                    //cellValue = null;
                                                                    cellStringValue = "null";
                                                                    ins_cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                    break;

                                                            case ERROR:
                                                                    //cellValue = null;
                                                                    cellStringValue = "null";
                                                                    ins_cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                    break;
                                                            }
                                                    }
                                                    iRow = iExcelCell;
                                                    //paramindex++;
                                            }

                                            if(iRow==iExcelLastCell)
                                            {
                                                    paramindex++;
                                                    ins_cStmt.setInt(paramindex, iExcelRow);
                                                    conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + iExcelRow + "\n";
                                            }
                                            System.out.println(conCatValue);
                                            ins_cStmt.execute();
                                            ins_cStmt.close();
                                    }	
                                    fis.close();
                            }
                            catch(Exception ex)
                            {
                                    System.out.println("Next Record");
                            }			
                    }
                    cStmt.close();
                    conn.commit();
                    conn.close();

            } 
            catch (Exception ex) {
                    // TODO: handle exception
                    System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
            }
    }

    /*
    public static void SearchExcuteScript() throws Exception
    {
            SQLObj sqlObj = new SQLObj();
            Connection conn = null;
            CallableStatement cStmt = null;
            ResultSet rsTestCase = null;
            ResultSet rsScript = null;

            try
            {
                    conn = sqlObj.ConnToDB();
                    cStmt = conn.prepareCall("{call Search_TestCaseID_InTestCase()}");
                    rsTestCase = cStmt.executeQuery();

                    while(rsTestCase.next())
                    {
                            //TC.ID,
                            //TC.TESTCASEID,
                            //TC.TESTCASEDESCRIPTION,
                            //TC.ITERATION,
                            //TC.EXECUTEFLAG,
                            //TC.RESULTFLAG,
                            //TC.TEST_CYCLE,
                            //TC.TEST_RUN

                            //String id = rsTestCase.getNString("ID");
                            String testcaseid = rsTestCase.getNString("TESTCASEID");
                            String testcasedesc = rsTestCase.getNString("TESTCASEDESCRIPTION");
                            String iteration = rsTestCase.getNString("ITERATION");
                            String executeflag = rsTestCase.getNString("EXECUTEFLAG");
                            String resultflag = rsTestCase.getNString("RESULTFLAG");
                            int test_cycle = rsTestCase.getInt("TEST_CYCLE");
                            int test_run = rsTestCase.getInt("TEST_RUN");


                            if(executeflag.equals("Y"))
                            {
                                    cStmt = conn.prepareCall("{call Select_CreateTestScript(?)}");
                                    cStmt.setString(1, testcaseid);

                                    rsScript = cStmt.executeQuery();

                                    boolean execresult = true;

                                    while(rsScript.next())
                                    {
                                            //TS.FUNC_CODE,
                                            //TS.FUNC_MAP,
                                            //TS.PAGE_MODULE,
                                            //TS.PAGE_FIELD,
                                            //TS.ELEMENT_ID,
                                            //TS.ELEMENT_XPATH,
                                            //TS.PAGE_TITLE,
                                            //TS.ELEMENT_TYPE,
                                            //TS.ELEMENT_VALUE,
                                            //TS.ELEMENT_ACTION,
                                            //TS.AD_SEQUENCE,
                                            //TS.REF
                                            String func_code = rsScript.getNString("FUNC_CODE");
                                            String func_map = rsScript.getNString("FUNC_MAP");
                                            String page_module = rsScript.getNString("PAGE_MODULE");
                                            String page_field = rsScript.getNString("PAGE_FIELD");
                                            String element_id = rsScript.getNString("ELEMENT_ID");
                                            String element_xpath = rsScript.getNString("ELEMENT_XPATH");
                                            String page_title = rsScript.getNString("PAGE_TITLE");
                                            String element_type = rsScript.getNString("ELEMENT_TYPE");
                                            String element_value = rsScript.getNString("ELEMENT_VALUE");
                                            String element_action = rsScript.getNString("ELEMENT_ACTION");
                                            String field_reference = "";

                                            //MapElementType(String Module, String Field, String Title, String Elementtype, String Elementid, String Elementxpath, String Elementvalue, String Action, String ReferenceField, boolean TakeScreenshot, String testcaseid)
                                            execresult = MapElementType(page_module, page_field, page_title, element_type, element_id, element_xpath, element_value, element_action, "", false, testcaseid);
                                    }
                            }

                    }


            }
            catch(Exception ex)
            {
                    System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
            }


    }
    */
    
    public static void readAppDatasource_TestData(String filename, String spcommand) throws Exception
    {
            SQLObj sqlObj = new SQLObj();

            try
            {
                    Connection conn = sqlObj.ConnToDB();
                    CallableStatement cStmt = null;
                    ResultSet rs = null;

                    //String filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TestDataTemplate.xlsx";

                    cStmt = conn.prepareCall("{call Search_All_InAppDataSource()}");
                    //cStmt = conn.prepareCall(spcommand);
                    rs = cStmt.executeQuery();

                    while(rs.next())
                    {
                            String testcaseid = rs.getNString("TESTCASEID");
                            String usercode = rs.getNString("USER_CODE");
                            String func_code = rs.getNString("FUNCTION_CODE");

                            FileInputStream fis = new FileInputStream(filename);
                            XSSFWorkbook workbook = new XSSFWorkbook(fis);

                            try
                            {
                                    XSSFSheet worksheet = workbook.getSheet(func_code);

                                    int iLastRow = worksheet.getLastRowNum();
                                    int iLastCell = worksheet.getRow(0).getLastCellNum() - 1;

                                    for(int iCell=1; iCell<=iLastCell; iCell++)
                                    {
                                            String conCatValue = "";
                                            String strFieldId = "";
                                            int paramindex = 5;

                                            for(int iRow=1; iRow<=iLastRow; iRow++)
                                            {
                                                    cStmt = conn.prepareCall(spcommand);
                                                    //TESTCASEID
                                                    cStmt.setString(1, testcaseid);
                                                    conCatValue = 1 + "|" + "STRING" + "-" + testcaseid + "\n";

                                                    //USER_CODE
                                                    cStmt.setString(2, usercode);
                                                    conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + usercode + "\n";

                                                    //FUNCTION_CODE
                                                    cStmt.setString(3, func_code);
                                                    conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + func_code + "\n";

                                                    //FIELD_ID
                                                    strFieldId = worksheet.getRow(iRow).getCell(0).getStringCellValue();
                                                    cStmt.setString(4, strFieldId);
                                                    conCatValue = conCatValue + 3 + "|" + "STRING" + "-" + strFieldId + "\n";

                                                    XSSFCell cellObj = worksheet.getRow(iRow).getCell(iCell);
                                                    String cellStringValue = "";
                                                    int cellIntValue = 0;

                                                    //FIELD_VALUE
                                                    if(cellObj == null)
                                                    {
                                                            cellStringValue = cellObj.getStringCellValue();
                                                            cStmt.setString(paramindex, cellStringValue);
                                                    }
                                                    else
                                                    {	
                                                            switch(cellObj.getCellType())
                                                            {
                                                                    case STRING:
                                                                            //cellValue = "STRING" + "##" + cellObj.getStringCellValue();
                                                                            cellStringValue = cellObj.getStringCellValue();
                                                                            cStmt.setString(paramindex, cellStringValue);
                                                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                            break;

                                                                    case NUMERIC:
                                                                            //cellValue = "INT" + "##" + NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                                            cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                                            cStmt.setInt(paramindex, cellIntValue);
                                                                            conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellStringValue + "\n";
                                                                            break;

                                                                    case BLANK:
                                                                            //cellValue = "STRING" + "##" + null;
                                                                            cellStringValue = "null";
                                                                            cStmt.setString(paramindex, cellStringValue);
                                                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                            break;

                                                                    case ERROR:
                                                                            //cellValue = "STRING" + "##" + null;
                                                                            cellStringValue = "null";
                                                                            cStmt.setString(paramindex, cellStringValue);
                                                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                            break;
                                                            }
                                                    }

                                                    cStmt.setInt(6, iCell);
                                                    conCatValue = conCatValue + 6 + "|" + "INT" + "-" + iCell + "\n";

                                                    System.out.println(conCatValue);
                                                    cStmt.execute();
                                                    cStmt.close();

                                            }

                                    }
                            }
                            catch(Exception ex)
                            {
                                    System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
                                    System.out.println("Next record");
                            }
                            fis.close();
                    }
                    conn.commit();
                    conn.close();
            }
            catch(Exception ex)
            {
                    System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" + "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
            }
    }


    //*********************************************************************************************************************
    /*
    public static void readMenu(String testcaseid, String MainMenumodule, String user) throws IOException
    {
            Properties propObj = new Properties();
            ArrayList<String> arrUnivTellerMenu = new ArrayList();
            arrUnivTellerMenu = propObj.getArrTellerMenu();

            int iArrCount = arrUnivTellerMenu.size() - 1;
            for(int iArr=0; iArr<=iArrCount; iArr++)
            {
                    String cellValue = arrUnivTellerMenu.get(iArr);
                    Scanner scanner = new Scanner(cellValue);
                    scanner.useDelimiter("&&");
                    int iScanner = 0;
                    while(scanner.hasNext())
                    {
                            String givenValue = scanner.next();
                            propObj.UnivTellerMenu(iScanner, givenValue);
                            iScanner++;
                    }

                    String module = propObj.getMenumodule();
                    String worksheetcode_main = propObj.getMenuWorksheetCodeMain();
                    String worksheetcode_sub = propObj.getMenuWorksheetCodeSub();
                    String worksheetname = propObj.getMenuWorksheet();
                    String submenu = propObj.getDataSourceSubMenu();

                    if(module.equalsIgnoreCase(MainMenumodule))
                    {	
                            String filename = "";
                            String worksheet = "";
                            switch(user)
                            {
                            case "Maker":
                                    filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\UniversalTellerMenu.xlsx";
                                    worksheet = "MenuList";
                                    getMenu_Main(filename, worksheet);
                                    readMenu_Main(module);
                                    getMenu_Sub(filename, worksheetcode_main, worksheetcode_sub, worksheetname);
                                    readMenu_Sub(testcaseid, submenu);		
                                    break;

                            case "Authoriser":
                                filename = "D:\\\\QA\\\\Projects\\\\Test Automation\\\\Java_SeleniumWebDriver_Demo\\\\templates\\\\ServiceManagerMenu.xlsx";
                                    worksheet = "MenuList";
                                    getMenu_Main(filename, worksheet);
                                    readMenu_Main(module);
                                    getMenu_Sub(filename, worksheetcode_main, worksheetcode_sub, worksheetname);
                                    readMenu_Sub(testcaseid, submenu);
                                    break;
                            }
                    }
            }
    }
    */
    //*********************************************************************************************************************

    //*********************************************************************************************************************
    /*
    public static void getMenuList(String filename, String worksheetname) throws IOException
    {
            Properties propObj = new Properties();
            //ArrayList<String> arrUnivTellerMenuMain = new ArrayList<String>();

            FileInputStream fis = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet worksheet = workbook.getSheet(worksheetname);
            int iExcelLastRow = worksheet.getLastRowNum() - 1;

            for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
            {
                    String conCatValue = "";
                    for(int iExcelCell=0; iExcelCell<=14; iExcelCell++)
                    {
                            XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                            String cellValue = "";
                            if(cellObj == null)
                            {
                                    cellValue = null;
                            }
                            else
                            {
                                    switch(cellObj.getCellType())
                                    {
                                    case STRING:
                                            cellValue = cellObj.getStringCellValue();
                                            break;

                                    case NUMERIC:
                                            cellValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                            break;

                                    case BLANK:
                                            cellValue = null;
                                            break;
                                    }
                            }

                            conCatValue = conCatValue + cellValue + "&&";
                    }
                    //arrUnivTellerMenuMain.add(conCatValue);
            }
            //propObj.setArrTellerMenuMain(arrUnivTellerMenuMain);
            fis.close();
    }

    public static void readMenu_Main(String module) throws IOException
    {
            Properties propObj = new Properties();
            ArrayList<String> arrUnivMainMenu = new ArrayList<String>();
            ArrayList<String> arrScript = new ArrayList();

            arrUnivMainMenu = propObj.getArrTellerMenuMain();
            arrScript = propObj.getArrScript();

            int iScriptCount = 0;
            if(arrScript == null)
            {
                    iScriptCount = 0;
            }
            else
            {
                    iScriptCount = arrScript.size();
            }

            int iArrCount = arrUnivMainMenu.size() - 1;
            for(String strValue: arrUnivMainMenu)
            {
                    if(strValue.contains(module))
                    {
                            if(arrScript == null)
                            {
                                    arrScript = new ArrayList();
                                    arrScript.add(iScriptCount, strValue);					
                            }
                            else
                            {
                                    //iScriptCount++;
                                    arrScript.add(iScriptCount, strValue);
                                    break;
                            }
                    }
            }
            propObj.setArrScript(arrScript);
    }

    public static void getMenu_Sub(String filename, String worksheetcodemain, String worksheetcodesub, String worksheetname) throws IOException
    {
            Properties propObj = new Properties();
            ArrayList<String> arrUnivSubMenu = new ArrayList();

            FileInputStream fis = new FileInputStream(filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet worksheet = workbook.getSheet(worksheetcodemain);
            int iExcelLastRow = worksheet.getLastRowNum();
            //String submenu = propObj.getDataSourceSubMenu();

            for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
            {
                    String conCatValue = "";
                    for(int iExcelCell=0; iExcelCell<=14; iExcelCell++)
                    {
                            XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                            String cellValue = "";
                            if(cellObj == null)
                            {
                                    cellValue = null;
                            }
                            else
                            {
                                    switch(cellObj.getCellType())
                                    {
                                    case STRING:
                                            cellValue = cellObj.getStringCellValue();
                                            break;

                                    case NUMERIC:
                                            cellValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                            break;

                                    case BLANK:
                                            cellValue = null;
                                            break;
                                    }
                            }

                            conCatValue = conCatValue + cellValue + "&&";
                    }
                    arrUnivSubMenu.add(conCatValue);
            }
            propObj.setArrTellerMenuSub(arrUnivSubMenu);
            fis.close();
    }	

    public static void readMenu_Sub(String testcaseid, String submenu) throws IOException
    {
            Properties propObj = new Properties();
            ArrayList<String> arrUnivSubMenu = new ArrayList();
            ArrayList<String> arrScript = new ArrayList();

            arrUnivSubMenu = propObj.getArrTellerMenuSub();
            arrScript = propObj.getArrScript();
            int iScriptCount = 0;

            if(arrScript == null)
            {
                    iScriptCount = 0;
            }
            else
            {
                    iScriptCount = arrScript.size();
            }

            //System.out.println("Array List before TellerMenu_Sub \n" + arrScript);
            int iArrCount = arrUnivSubMenu.size() - 1;
            for(String strValue: arrUnivSubMenu)
            {
                    if(strValue.contains(submenu))
                    {
                            if(arrScript == null)
                            {
                                    arrScript = new ArrayList();
                                    arrScript.add(iScriptCount, strValue);
                            }
                            else
                            {
                                    //iScriptCount++;
                                    arrScript.add(iScriptCount, strValue);
                                    break;
                            }
                    }
            }
            //System.out.println("Array List after TellerMenu_Sub \n" + arrScript);
            propObj.setArrScript(arrScript);
    }

    public static void ExecuteArrList() throws Exception
    {
            Properties propObj = new Properties();
            SeleniumObj seleniumObj = new SeleniumObj();

            String browsername = propObj.getBrowser();
            String baseurl = propObj.getBaseUrl();

            //seleniumObj.Loginto(browsername, baseurl);

            ArrayList<String> arrScripts = new ArrayList();
            ArrayList<String> arrTestData = new ArrayList();
            ArrayList<String> arrMapTD = new ArrayList();

            int iArrCount = 0;
            int iArrCountTD = 0; 
            int iCtr = 0;

            //arrTestData = propObj.getArrTestData();
            //arrMapTD = propObj.getArrMapTestData();

            if(arrScripts == null || arrScripts.isEmpty())
            {
                    arrScripts = propObj.getArrScript();
                    iArrCount = arrScripts.size() - 1;

                    for(int iArr=0; iArr<=iArrCount; iArr++)
                    {	
                            String begin = "";
                            String testcaseid = "";
                            String module = "";
                            String field = "";
                            String elementid = "";
                            String elementxpath = "";
                            String linktext = "";
                            String pageurl = "";
                            String title = "";
                            String elementtype = "";
                            String value = "";
                            String action = "";
                            String menureference = "";
                            String worksheetreferenc = "";
                            String fieldreference = "";
                            String message = "";
                            String result = "";

                            String arrValue = arrScripts.get(iArr);
                            Scanner scanner = new Scanner(arrValue);
                            scanner.useDelimiter("&&");

                            if(arrValue.contains("Begin"))
                            {
                                    begin = scanner.next();
                                    if(begin.equalsIgnoreCase("Begin"))
                                    {
                                            iCtr = 1;
                                            seleniumObj.Loginto(browsername, baseurl);
                                            testcaseid = scanner.next();		
                                            if(!testcaseid.equals("null"))
                                            {
                                                    propObj.setCurrentTestCaseId(testcaseid);
                                            }
                                    }
                            }
                            else if(arrValue.equalsIgnoreCase("End"))
                            {
                                    iCtr = 0;
                                    seleniumObj.CloseBrowser();
                            }
                            else
                            {
                                    int iScan = 0;
                                    while(scanner.hasNext())
                                    {
                                            String givenValue = scanner.next();
                                            propObj.MapScript(iScan, givenValue);
                                            iScan++;
                                    }
                                     module = propObj.getScriptModule();
                                     field = propObj.getScriptField();
                                     elementid = propObj.getScriptElementId();
                                     elementxpath = propObj.getScriptElementXpath();
                                     linktext = propObj.getScriptLinkText();
                                     pageurl = propObj.getScriptPageUrl();
                                     title = propObj.getScriptTitle();
                                     elementtype = propObj.getScriptElementType();
                                     value = propObj.getScriptValue();
                                     action = propObj.getScriptAction();
                                     menureference = propObj.getScriptMenuReference();
                                     worksheetreferenc = propObj.getScriptWorksheetReference();
                                     fieldreference = propObj.getScriptFieldReference();
                                     message = propObj.getScriptMessage();
                                     result = propObj.getScriptResult();
                                     boolean TakeScreenshot = false;

                                     if(result.equalsIgnoreCase("TakeScreenshot"))
                                     {
                                             TakeScreenshot = true;
                                             testcaseid = propObj.getCurrentTestCaseId();
                                     }

                                     if(!linktext.equals("null"))
                                     {
                                             String givenValue = "";
                                             //givenValue = readTestData(linktext, field);
                                             value = givenValue;
                                     }

                                    if(MapElementType(module, field, title, elementtype, elementid, elementxpath, value, action, fieldreference, TakeScreenshot, testcaseid))
                                    {
                                            System.out.println(iCtr + " | " + module + " | " + "True" + " | " + "FIELD :" + field + " | " + "ELEMENT ID :" + elementid + " | " + "ELEMENT XPATH :" + elementxpath + " | " + "ELEMENT TYPE :" + elementtype + " | " + "ACTON :" + action);
                                    }
                                    else
                                    {
                                            System.out.println(iCtr + " | " + module + " | " + "False" + " | " + "FIELD :" + field + " | " + "ELEMENT ID :" + elementid + " | " + "ELEMENT XPATH :" + elementxpath + " | " + "ELEMENT TYPE :" + elementtype + " | " + "ACTON :" + action);
                                    }
                                    iCtr++;
                            }
                    }
            }
    }
    */
    //*********************************************************************************************************************

    //*********************************************************************************************************************
    /*
    public static void StoreToTemp(String value, String fieldreference) throws Exception
    {
            Properties propObj = new Properties();
            ArrayList<String> arrTempStorage = new ArrayList<String>();
            int iArrCount = 0;
            String givenValue = "";
            String scanValue = "";

            arrTempStorage = propObj.getArrTempStorage();
            givenValue = fieldreference + "&&" + value;

            if(arrTempStorage == null)
            {
                    arrTempStorage = new ArrayList<String>();
                    arrTempStorage.add(iArrCount, givenValue);
            }
            else
            {
                    //Check if ArrayList Temp Storage is with existing record
                    boolean editFlag = false;

                    for(String strValue: arrTempStorage)
                    {
                            if(strValue.contains(fieldreference))
                            {
                                    Scanner scanner = new Scanner(strValue);
                                    scanner.useDelimiter("&&");
                                    //int iScan = scanner.;
                                    for(int iScanCtr=0; iScanCtr<=2; iScanCtr++)
                                    {
                                            scanValue = new String();
                                            scanValue = scanner.next();
                                            if(!scanValue.equals(fieldreference))
                                            {
                                                    scanValue = scanValue + "*" + value;
                                                    editFlag = true;
                                                    break;
                                            }
                                    }
                            }
                            if(editFlag)
                            {
                                    break;
                            }
                            iArrCount++;
                    }

                    if(editFlag)
                    {
                            givenValue = "";
                            givenValue = fieldreference + "&&" + scanValue;
                            arrTempStorage.set(iArrCount, givenValue);
                            System.out.println(arrTempStorage);
                    }
                    else
                    {
                            iArrCount = arrTempStorage.size();
                            arrTempStorage.add(iArrCount, givenValue);				
                    }

            }

            propObj.setArrTempStorage(arrTempStorage);
    }
    */
    //*********************************************************************************************************************

    //*********************************************************************************************************************
    /*
    public static String ReadToTemp(String fieldreference) throws Exception
    {
            Properties propObj = new Properties();
            ArrayList<String> arrTempStorage = new ArrayList<String>();
            int iArrCount = 0;
            String value = "";

            arrTempStorage = propObj.getArrTempStorage();
            for(String stored: arrTempStorage)
            {
                    if(stored.contains(fieldreference))
                    {
                            Scanner scanner = new Scanner(stored);
                            int iScan = 0;
                            scanner.useDelimiter("&&");
                            while(scanner.hasNext())
                            {
                                    if(iScan==1)
                                    {
                                            value = scanner.next();						
                                    }
                                    else
                                    {
                                            scanner.next();
                                    }
                                    iScan++;
                            }
                    }
            }

            return value;
    }
    */
    //*********************************************************************************************************************
    
    public static void ClearArrTempStorage(){
        Properties propObj = new Properties();
        ArrayList<String> arrTempStorage = new ArrayList<String>();
        arrTempStorage = propObj.getArrTempStorage();
        if(arrTempStorage != null){
            int arrcount = arrTempStorage.size();
            if(arrcount > 0){
                arrTempStorage.clear();
            }            
        }

    }
    
    //*********************************************************************************************************************


    //*********************************************************************************************************************

    public static void TakeScreenShot(WebDriver webdriver, String testcaseid, String module, String field) throws IOException
    {
            try
            {
                    TimeUnit.SECONDS.sleep(15);

                    Date dateNow = new Date();
                    SimpleDateFormat fdate = new SimpleDateFormat("E yyyy.MM.dd"+"_"+"hh.mm.ss");

                    String filename = testcaseid + "_" + module + "_" + field + "_" + fdate.format(dateNow) + ".png";
                    //String fileWithPath = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\screenshots\\" + filename;
                    String fileWithPath = "D:\\temp\\screenshot\\" + filename;
                    TakesScreenshot scrnshot = ((TakesScreenshot)webdriver);
                    File SrcFile = scrnshot.getScreenshotAs(OutputType.FILE);
                    File DestFile = new File(fileWithPath);
                    FileUtils.copyFile(SrcFile, DestFile);			
            }
            catch(Exception ex)
            {
                    System.out.println("Error on Screenshot - " + ex.getMessage());
            }

            //FileUtils fu = new FileUtils();

    }

    //*********************************************************************************************************************

    /* 
    public static void getTestData(String filename) throws IOException
    {
            String strSheetName = "";
            try
            {
                    Properties propObj = new Properties();
                    FileInputStream fis = new FileInputStream(filename);
                    XSSFWorkbook workbook = new XSSFWorkbook(fis);
                    int iCtr = 0;
                    int iSheet = 0;
                    iSheet = workbook.getNumberOfSheets();

                    for(int iSheetNum = 0; iSheetNum <= iSheet; iSheetNum++)
                    {
                            strSheetName = workbook.getSheetName(iSheetNum);

                            if(workbook.getSheet(strSheetName) != null)
                            {
                                    XSSFSheet worksheet = workbook.getSheet(strSheetName);
                                    int intLastRow = worksheet.getLastRowNum();
                                    int intLastCell = worksheet.getRow(0).getLastCellNum() - 1;
                                    String strField = "";
                                    String strGivenvalue = "";
                                    String strFieldGivenvalue = "";
                                    String strGroupIter = "";
                                    String strIteration = "";
                                    ArrayList<String> arrTestData = new ArrayList();

                                    arrTestData = propObj.getArrTestData();

                                    for(int iCell = 1; iCell<=intLastCell; iCell++)
                                    {
                                            strGroupIter = "";
                                            for(int iRow = 1; iRow<=intLastRow; iRow++)
                                            {
                                                    strGivenvalue = "";
                                                    XSSFCell cellObj = worksheet.getRow(iRow).getCell(iCell);
                                                    if(cellObj == null)
                                                    {
                                                            strGivenvalue = null;
                                                    }
                                                    else
                                                    {
                                                            strField = worksheet.getRow(iRow).getCell(0).getStringCellValue();
                                                            switch(cellObj.getCellType())
                                                            {
                                                            case STRING:
                                                                    strGivenvalue = cellObj.getStringCellValue();
                                                                    break;

                                                            case NUMERIC:
                                                                    strGivenvalue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                                    break;

                                                            case BLANK:
                                                                    strGivenvalue = null;
                                                                    break;
                                                            }
                                                    }

                                                    strFieldGivenvalue = strField + "##" + strGivenvalue;
                                                    if(strGroupIter == "")
                                                    {
                                                            strGroupIter = strFieldGivenvalue;
                                                    }
                                                    else
                                                    {
                                                            strGroupIter = strGroupIter + "&&" + strFieldGivenvalue ;							
                                                    }
                                            }
                                            //strIteration = strIteration + strGroupIter + "&&";
                                            strGroupIter = strSheetName + "!!" + strGroupIter;
                                            if(arrTestData ==  null)
                                            {
                                                    arrTestData = new ArrayList();
                                                    arrTestData.add(iCtr, strGroupIter);		
                                            }
                                            else
                                            {
                                                    arrTestData.add(iCtr, strGroupIter);
                                            }

                                            iCtr++;
                                    }
                                    System.out.println(arrTestData);
                                    propObj.setArrTestData(arrTestData);
                            }

                    }

            }
            catch(Exception ex)
            {
                    System.out.println("Worksheet does not exist :" + strSheetName + "/n" + "Error messessage : " + ex.getLocalizedMessage());
            }
    }

    public static String readTestData(int iteration, String menucode, String field) throws IOException
    {
            Properties propObj = new Properties();
            ArrayList<String> arrTestData = new ArrayList();
            boolean flag = false;

            arrTestData = propObj.getArrTestData();
            int iTDSize = arrTestData.size();
            String givenValue = "";

            for(String strValue : arrTestData)
            {
                    if(strValue.contains(menucode))
                    {
                            String arrValue = "";
                            System.out.println(strValue);

                            Scanner scannerArr = new Scanner(strValue);
                            scannerArr.useDelimiter("!!");
                            while(scannerArr.hasNext())
                            {
                                    String strFieldGivenvalue = scannerArr.next();
                                    System.out.println(strFieldGivenvalue);
                                    if(!strFieldGivenvalue.equals(menucode))
                                    {
                                            Scanner scannerFieldGivenvalue = new Scanner(strFieldGivenvalue);
                                            scannerFieldGivenvalue.useDelimiter("&&");
                                            while(scannerFieldGivenvalue.hasNext())
                                            {
                                                    String value = "";
                                                    givenValue = "";
                                                    value = scannerFieldGivenvalue.next();
                                                    //System.out.println(value);
                                                    if(value.contains(field))
                                                    {
                                                            //givenValue = scannerFieldGivenvalue.next();
                                                            //System.out.println(givenValue);

                                                            Scanner scanner = new Scanner(value);
                                                            scanner.useDelimiter("##");
                                                            while(scanner.hasNext())
                                                            {
                                                                    String checkValue = "";
                                                                    String UseValue = "";
                                                                    checkValue = scanner.next();
                                                                    if(checkValue.equals(field))
                                                                    {

                                                                            UseValue = scanner.next();
                                                                            givenValue = UseValue;
                                                                            flag = true;

                                                                            System.out.println("Match Field :" + checkValue + " | " + flag);
                                                                            System.out.println("Use Value :" + UseValue);
                                                                            break;
                                                                    }
                                                                    else
                                                                    {
                                                                            UseValue = scanner.next();
                                                                            flag = false;

                                                                            System.out.println("Unmatch Field :" + checkValue + " | " + flag);
                                                                            System.out.println("Use Value :" + UseValue);
                                                                            break;
                                                                    }
                                                            }
                                                    }
                                                    if(flag)
                                                    {
                                                            break;
                                                    }
                                            }								
                                    }

                                    if(flag)
                                    {
                                            break;
                                    }

                            }
                    }
                    if(flag)
                    {
                            break;
                    }
            }

            return givenValue;
    }
    */
    //*********************************************************************************************************************

    public static int getArrayMenucodeTestdata(String menucode) throws Exception
    {
            int iCount = 0;
            int iCtr = 0;
            int iFlag = 0;

            ArrayList<String> arrTestData = new ArrayList();
            ArrayList<String> arrMapTD = new ArrayList();
            Properties propObj = new Properties();

            arrTestData = propObj.getArrTestData();

            for(String value: arrTestData)
            {
                    if(value.contains(menucode))
                    {
                            arrMapTD.add(iCtr, NumberToTextConverter.toText(iFlag));
                            iCtr++;
                    }
                    iFlag++;
            }
            iCount = iCtr;

            propObj.setArrMapTestData(arrMapTD);
            return iCount;
    }

    public static CallableStatement statement(String calltype, int Index, String Value)
    {
            CallableStatement cStmt = null; 
            switch(calltype)
            {
            case "Int":
                    break;

            case "String":
                    break;

            case "Boolean":
                    break;
            }

            return cStmt;

    }

    //***********************************************************************************************************************************    
}
