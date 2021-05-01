package ewb.qa.tdd;

import static com.sun.glass.ui.Cursor.setVisible;
import ewb.qa.tdd.GUI.LoginGUI;
import java.util.ArrayList;

public class main {
    
public static ExcelObj excelObj;
public static SeleniumObj seleniumObj;
public static Properties propObj;
public static LoginGUI loginGui;
private static String globalVersionRelease;
    
    public static void main(String[] args)
    {
        try
            {
                String filename = "";
                String worksheetname = "";
                String usercode = "";
                int lastcolumn = 0;
                String spcommand = "";
                String qcommand = "";
                String sequenceflag = "";

                
                seleniumObj = new SeleniumObj();
                excelObj = new ExcelObj();
                propObj = new Properties();

                
                setVisible(false);
                LoginGUI form = new LoginGUI();
                form.setVisible(true);
                
                
                //ArrayList<String> arrArchTestCases = new ArrayList();

                //***************************************************************************************************************
                //Consolidate and Insert to dbo.TESTCASES
                /*
                filename = new String("D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\T24 Architecture Layout.xlsx");
                worksheetname = new String("TestCases");
                excelObj.getArchTestCases(filename, worksheetname);
                System.out.println("Successfully Insert data to dbo.TESTCASES");
                */
                //***************************************************************************************************************

                //***************************************************************************************************************
                //Consolidate and Insert to dbo.APP_DATASOURCE
                /*
                filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\T24 Architecture Layout.xlsx";
                worksheetname = "Automation App DataSource";
                lastcolumn = 14;
                spcommand = "{call Insert_AppDataSource(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                qcommand = "TRUNCATE TABLE dbo.APP_DATASOURCE";
                excelObj.getArchDataSource(filename, worksheetname, spcommand, qcommand, lastcolumn);
                //excelObj.getSpecificWorksheet_Excel(filename, worksheetname, lastcolumn, spcommand, qcommand);
                System.out.println("Successfully Insert data to dbo.APP_DATASOURCE");
                */
                //***************************************************************************************************************

                //***************************************************************************************************************
                //Consolidate and Insert to dbo.PAGE_TESTSCRIPT - MAKER
                /*
                filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TDD_FunctionalityMaker.xlsx";
                spcommand = "{call Insert_PageTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                lastcolumn = 14;
                worksheetname = "";
                //getGetWorksheetList_Excel(String filename, String spcommand, int lastcolumn)
                excelObj.getGetWorksheetList_Excel(filename, worksheetname, spcommand, lastcolumn);
                System.out.println("Successfully Insert data to dbo.PAGE_TESTSCRIPT - MAKER");
                */
                //***************************************************************************************************************

                //readArchDataSource(testcaseid);

                //***************************************************************************************************************
                //Consolidate and Insert to dbo.PAGE_TESTSCRIPT - AUTHORISER
                /*
                filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TDD_FunctionalityAuthoriser.xlsx";
                spcommand = "{call Insert_PageTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                lastcolumn = 14;
                worksheetname = "";
                //getGetWorksheetList_Excel(String filename, String spcommand, int lastcolumn)
                excelObj.getGetWorksheetList_Excel(filename, worksheetname, spcommand, lastcolumn);
                System.out.println("Successfully Insert data to dbo.PAGE_TESTSCRIPT - AUTHORISER");
                */
                //***************************************************************************************************************

                //***************************************************************************************************************
                //Consolidate and Insert to dbo.USER_MENUMAP - MAKER
                /*
                filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\UniversalTellerMenu.xlsx";
                worksheetname = "User0001";
                spcommand = "{call Insert_Menumap(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                lastcolumn = 17;
                qcommand = "";
                sequenceflag = "N";
                //getSpecificWorksheet_Excel(String filename, String worksheetname, int lastcolumn, String spcommand, String qcommand, String sequenceflag)
                excelObj.getSpecificWorksheet_Excel(filename, worksheetname, lastcolumn, spcommand, qcommand, sequenceflag);
                System.out.println("Successfully Insert data to dbo.USER_MENUMAP - MAKER");
                */
                //***************************************************************************************************************

                //***************************************************************************************************************
                //Consolidate and Insert to dbo.USER_MENUMAP - AUTHORISER
                /*
                filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\ServiceManagerMenu.xlsx";
                worksheetname = "User0002";
                lastcolumn = 17;
                spcommand = "{call Insert_Menumap(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                qcommand = "";
                sequenceflag = "N";
                excelObj.getSpecificWorksheet_Excel(filename, worksheetname, lastcolumn, spcommand, qcommand, sequenceflag);
                //getSpecificWorksheet_Excel(String filename, String worksheetname, int lastcolumn, String command)
                System.out.println("Successfully Insert data to dbo.USER_MENUMAP - AUTHORISER");
                */
                //***************************************************************************************************************

                //***************************************************************************************************************
                /*
                filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\TestDataTemplate.xlsx";
                spcommand = "{call Insert_Testdata(?,?,?,?,?,?)}";

                excelObj.readAppDatasource_TestData(filename, spcommand);
                System.out.println("Successfully Insert data to dbo.TESTDATA");
                */
                //***************************************************************************************************************			


                //***************************************************************************************************************
                /*
                excelObj.readUserMenuMap();
                System.out.println("Successfully Insert data to dbo.MENU_TESTSCRIPT");
                */
                //***************************************************************************************************************
                //excelObj.readArchTestCases();


                //***************************************************************************************************************
                /*
                System.out.println("Executing Script............");
                String browsername = propObj.getBrowser();
                String baseurl = propObj.getBaseUrl();

                seleniumObj.Loginto(browsername, baseurl);
                excelObj.SearchExcuteScript();
                */
                //***************************************************************************************************************

                //excelObj.ExecuteArrList();
                //seleniumObj.CloseBrowser();

        }
        catch(Exception ex)
        {
                System.out.println(ex.getMessage().toString());
                //System.out.println(ex.getLocalizedMessage());

        }
    }
    

}
