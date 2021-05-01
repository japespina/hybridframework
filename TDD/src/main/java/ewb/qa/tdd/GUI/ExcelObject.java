/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ewb.qa.tdd.GUI;
//import static ewb.qa.tdd.ExcelObj.ReadToTemp;
//import static ewb.qa.tdd.ExcelObj.StoreToTemp;
import static ewb.qa.tdd.ExcelObj.TakeScreenShot;
import ewb.qa.tdd.SQLObj;
import ewb.qa.tdd.SeleniumObj;
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
import java.util.Date;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

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
import java.text.DateFormat;
import static java.time.LocalDateTime.now;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import java.text.NumberFormat;
/**
 *
 * @author JPE61800
 */
public class ExcelObject {
    static WebDriver driver;
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
    public ewb.qa.tdd.Properties propertiesObj;
        
    public boolean getSpecificWorksheet_Excel(String filename, String worksheetname, String spcommand) throws IOException{
            JFrame jFrame = new JFrame();
            ewb.qa.tdd.Properties propObj = new ewb.qa.tdd.Properties();
            SQLObj sqlObj = new SQLObj();
            Connection conn = null;
            CallableStatement cStmt = null;
            Statement Stmt = null;
            int indexRow = 0;
            int indexCell = 0;
            //int paramindex = 0;
            //ArrayList<String> arrUnivTellerMenu = new ArrayList();
            boolean result = false;
            FileInputStream fis = null;
            String strValue = "";
            String conCatValue = "";
            
            try
            {	
                    result = true;
                    fis = new FileInputStream(filename);
                    XSSFWorkbook workbook = new XSSFWorkbook(fis);
                    XSSFSheet worksheet = workbook.getSheet(worksheetname);
                    int iExcelLastRow = worksheet.getLastRowNum();
                    int iExcelLastCell = worksheet.getRow(0).getLastCellNum() -1;
                    String projcode = MainGUI.getProjectCode();
                    
                    conn = sqlObj.ConnToDB();
                    for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                    {
                            indexRow = iExcelRow;

                            int iRow = 0;
                            int paramindex = 3;
                            //cStmt = conn.prepareCall("{call Insert_Menumap(?,?,?,?,?,?)}");
                            cStmt = conn.prepareCall(spcommand);

                            cStmt.setNString(1, projcode);
                            conCatValue = 1 + "|" + "STRING" + "-" + projcode + "\n";
                            cStmt.setString(2, worksheetname);
                            conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + worksheetname + "\n";

                            for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                            {	
                                    indexCell = iExcelCell;
                                    String cellStringValue = "";
                                    int cellIntValue = 0;

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
                                                    strValue = cellStringValue;
                                                    break;

                                            case NUMERIC:
                                                    //cellValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                    cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                    cStmt.setInt(paramindex, cellIntValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellIntValue + "\n";
                                                    strValue = NumberToTextConverter.toText(cellIntValue);
                                                    break;

                                            case BLANK:
                                                    //cellValue = null;
                                                    cellStringValue = "null";
                                                    cStmt.setString(paramindex, cellStringValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                    strValue = cellStringValue;
                                                    break;

                                            case ERROR:
                                                    cellStringValue = "null";
                                                    cStmt.setString(paramindex, cellStringValue);
                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                    strValue = cellStringValue;
                                                    break;
                                            }
                                    }
                                    iRow = iExcelCell;
                                    paramindex++;
                            }
                            System.out.println(conCatValue);
                            cStmt.execute();
                            cStmt.close();
                    }
                    fis.close();
                    conn.commit();
                    conn.close();
                    
                    return result;
                    
            }
            catch(Exception ex)
            {
                    result = false;
                    String errMessage = null;

                    errMessage = "File Path : " + filename + "\n" +
                        "Worksheet: " + worksheetname + ", " + "Row: " + indexRow + ", " + "Cell: " + indexCell + ", " + "Value: " + strValue + "\n" +
                        "Concate Value: " + conCatValue + "\n" + 
                        "Error Message: " + ex.getMessage() + 
                        "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                        "\n" + "Stack Trace: " + ex.getStackTrace() + 
                        "\n" + "Cause: " + ex.getCause();

                    System.out.println(errMessage);

                    JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                                   "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);
                    
                    fis.close();
                    return result;
            }
    }
                
    public void getExcelFieldMapping(String filename) throws IOException{
            JFrame jFrame = new JFrame();
            ewb.qa.tdd.Properties propObj = new ewb.qa.tdd.Properties();
            ewb.qa.tdd.SQLObj sqlObj = new ewb.qa.tdd.SQLObj();
            Connection conn = null;
            CallableStatement cStmt = null;
            Statement Stmt = null;

            try{
                    FileInputStream fis = new FileInputStream(filename);
                    XSSFWorkbook workbook = new XSSFWorkbook(fis);
                    XSSFSheet worksheet = null;
                    String worksheetname = "";

                    int countWorksheet = workbook.getNumberOfSheets() - 1;
                    int lastCell = 0;
                    int lastRow = 0;

                    conn = SQLObj.ConnToDB();
                    for(int indexWorksheet =0; indexWorksheet <=countWorksheet; indexWorksheet++){
                            worksheetname = workbook.getSheetName(indexWorksheet);
                            worksheet = workbook.getSheet(worksheetname);

                            lastRow = worksheet.getLastRowNum();
                            lastCell = worksheet.getRow(0).getLastCellNum() -1;
                            //int paramindex = 1;
                            //String conCatValue = "";
                            int countCell = 0;

                            for(int indexRow = 1; indexRow <= lastRow; indexRow++){       
                                    int paramindex = 1;
                                    String conCatValue = "";
                                    cStmt = conn.prepareCall("{call Insert_FieldMap(?,?,?,?,?,?,?,?,?)}");
                                    cStmt.setString(paramindex, worksheetname);
                                    conCatValue = paramindex + "|" + "STRING" + "-" + worksheetname + "\n";

                                    for(int indexCell = 0; indexCell <= lastCell; indexCell++){
                                            paramindex++;
                                            String cellStringValue = "";
                                            int cellIntValue = 0;

                                            XSSFCell cellObj = worksheet.getRow(indexRow).getCell(indexCell);
                                            if(cellObj == null){
                                                    cellStringValue = "null";
                                                    cStmt.setString(paramindex, cellStringValue);
                                            }
                                            else{
                                                    switch(cellObj.getCellType())
                                                    {
                                                    case STRING:
                                                            //cellValue = cellObj.getStringCellValue();
                                                            cellStringValue = cellObj.getStringCellValue();
                                                            if(indexCell == 0){
                                                                    cellStringValue = cellObj.getStringCellValue();
                                                                    CallableStatement cStmtSubmod = conn.prepareCall("{call Search_ProjectSubmodule_BySubname(?)}");
                                                                    cStmtSubmod.setNString(1, cellStringValue);
                                                                    ResultSet rsSubmod = cStmtSubmod.executeQuery();

                                                                    if(rsSubmod.next()){
                                                                            String id = NumberToTextConverter.toText(rsSubmod.getInt("ID"));
                                                                            String modcode = rsSubmod.getNString("MCODE");
                                                                            String subcode = rsSubmod.getNString("SCODE");
                                                                            String subname = rsSubmod.getNString("SNAME");

                                                                            //cellStringValue = modcode;
                                                                            cStmt.setString(paramindex, modcode);
                                                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + modcode + "\n";

                                                                            paramindex++;
                                                                            cStmt.setString(paramindex, subcode);
                                                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + subcode + "\n";
                                                                    }
                                                                    cStmtSubmod.close();
                                                            }
                                                            else{
                                                                    cStmt.setString(paramindex, cellStringValue);
                                                                    conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";                                                                            
                                                            }
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
                                            countCell = indexCell;
                                    }

                                    System.out.println(conCatValue);
                                    cStmt.execute();
                                    conn.commit();
                                    cStmt.close();
                            }
                    }
                    conn.close();

            }
            catch(Exception ex){
                    String errMessage = null;

                    errMessage = "Error Message: " + ex.getMessage() + 
                       "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                       "\n" + "Stack Trace: " + ex.getStackTrace() + 
                       "\n" + "Cause: " + ex.getCause();

                    System.out.println(errMessage);

                    JOptionPane.showMessageDialog(jFrame,  "Field Mapping Consolidation Failed " + 
                                   "\n" + errMessage, "Maintenance", 0);

            }

    }
                
    public static void getArchDataSource(String filename, String worksheetname, String spcommand, String qcommand) throws IOException{
        JFrame jFrame = new JFrame();
        ewb.qa.tdd.Properties propObj = new ewb.qa.tdd.Properties();
        //ArrayList<String> arrArchAutoAppDataSource = new ArrayList();
        FileInputStream fis = new FileInputStream(filename);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet worksheet = workbook.getSheet(worksheetname);
        int iExcelLastRow = worksheet.getLastRowNum() - 1;
        int iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;
        int indexRow = 0;
        int indexCell = 0;
        int indexTestCase = 0;
        
        SQLObj sqlObj = new SQLObj();
        Connection conn = null;
        CallableStatement cStmt = null;
        Statement Stmt = null;

        String globalProjCode = MainGUI.getProjectCode();
        
        try
        {
                conn = sqlObj.ConnToDB();
                Stmt = conn.createStatement();
                //int intStmt = Stmt.executeUpdate("TRUNCATE TABLE dbo.APP_DATASOURCE");
                int intStmt = Stmt.executeUpdate(qcommand);
                conn.commit();
                String TestCaseId = "";
                
                if(intStmt == -1)
                {
                        for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                        {
                                 indexRow = iExcelRow;
                                //cStmt = conn.prepareCall("{call Insert_AppDataSource(?,?,?,?,?,?,?,?,?,?)}");
                                cStmt = conn.prepareCall(spcommand);
                                String conCatValue = "";
                                int paramindex = 2;
                                int iRow = 0;
                                cStmt.setNString(1, globalProjCode);
                                conCatValue = 1 + "|" + "STRING" + "-" + globalProjCode + "\n";
                                
                                for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                {
                                            indexCell = iExcelCell;
                                            XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
                                            String cellStringValue = "";
                                            int cellIntValue = 0;

                                            if(cellObj == null)
                                            {
                                                    cellStringValue = cellObj.getStringCellValue();
                                                    cStmt.setString(paramindex, cellStringValue);
                                            }
                                            else{	
                                                        switch(cellObj.getCellType())
                                                        {
                                                                case STRING:
                                                                            //cellValue = "STRING" + "##" + cellObj.getStringCellValue();
                                                                            cellStringValue = cellObj.getStringCellValue();
                                                                            cStmt.setString(paramindex, cellStringValue);
                                                                            conCatValue = conCatValue + paramindex + "|" + "STRING" + "-" + cellStringValue + "\n";
                                                                            if(iExcelCell == 0){
                                                                                if(TestCaseId.equals("")){
                                                                                    TestCaseId = cellStringValue;
                                                                                    indexTestCase++;
                                                                                }
                                                                                else if(TestCaseId.equals(cellStringValue)){
                                                                                    indexTestCase++;
                                                                                }
                                                                                else if(!TestCaseId.equals(cellStringValue)){
                                                                                    TestCaseId = cellStringValue;
                                                                                    indexTestCase = 1;
                                                                                }
                                                                            }
                                                                            break;

                                                                case NUMERIC:
                                                                            //cellValue = "INT" + "##" + NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                                            //cellIntValue = Integer.parseInt(NumberToTextConverter.toText(cellObj.getNumericCellValue()));
                                                                            double dblValue = cellObj.getNumericCellValue();
                                                                            NumberFormat nf = NumberFormat.getNumberInstance();
                                                                            nf.setParseIntegerOnly(true);
                                                                            
                                                                            String strValue = nf.format(dblValue);
                                                                            cellIntValue = Integer.parseInt(strValue);
                                                                            cStmt.setInt(paramindex, cellIntValue);
                                                                            conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + cellIntValue + "\n";
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
                                        //cStmt.setInt(paramindex, iExcelRow);
                                        cStmt.setInt(paramindex, indexTestCase);
                                        conCatValue = conCatValue + paramindex + "|" + "INT" + "-" + iExcelRow + "\n";
                                }
                                /*
                                @proj_code NVARCHAR(20),
                                @testcaseid  NVARCHAR(12),
                                @usertype  NVARCHAR(10),
                                @user_code NVARCHAR(20),
                                @main_module  NVARCHAR(50),
                                @sub_module  NVARCHAR(50),
                                @function_code  NVARCHAR(20),
                                @function_map NVARCHAR(100),
                                @main_menu  NVARCHAR(100),
                                @main_menucode NVARCHAR(50),
                                @sub_menu  NVARCHAR(100),
                                @sub_menucode NVARCHAR(50),
                                @child_menu  NVARCHAR(100),
                                @child_menucode NVARCHAR(50),
                                @gchild1_menu NVARCHAR(100),
                                @gchild1_menucode NVARCHAR(50),
                                @gchild2_menu NVARCHAR(100),
                                @gchild2_menucode NVARCHAR(50),
                                @gchild3_menu NVARCHAR(100),
                                @gchild3_menucode NVARCHAR(50),
                                @nvartemp1 NVARCHAR(500),
                                @nvartemp2 NVARCHAR(500),
                                @nvartemp3 NVARCHAR(500),
                                @nvartemp4 NVARCHAR(500),
                                @nvartemp5 NVARCHAR(500),
                                @inttemp1 INT,
                                @inttemp2 INT,
                                @inttemp3 INT,
                                @inttemp4 INT,
                                @inttemp5 INT,
                                @sequence INT
                                */                                

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
                String errMessage = null;

                errMessage = "App-DataSource " + "\n" + 
                    "Worksheet: " + worksheetname + ", " + "Row: " + indexRow + ", " + "Cell: " + indexCell + "\n" +
                    "Error Message: " + ex.getMessage() + 
                    "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                    "\n" + "Stack Trace: " + ex.getStackTrace() + 
                    "\n" + "Cause: " + ex.getCause();

                System.out.println(errMessage);

                JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                               "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);
                fis.close();

        }

    }
                
    public static void getGetWorksheetList_Excel(String filename, String spcommand) throws IOException {
            JFrame jFrame = new JFrame();
            ewb.qa.tdd.Properties propObj = new ewb.qa.tdd.Properties();
            SQLObj sqlObj = new SQLObj();
            String projcode = MainGUI.getProjectCode();
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
            int indexRow = 0;
            int indexCell = 0;
            
            try
            {
                conn = sqlObj.ConnToDB();
                for(int iWorksheet=0; iWorksheet<=iWorkSheets; iWorksheet++)
                {
                        String conCatValue = "";
                        worksheetname = workbook.getSheetName(iWorksheet);			
                        worksheet = workbook.getSheet(worksheetname);

                        iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;
                        //if(iExcelLastCell == lastcolumn)
                        //{
                        //        System.out.println("Same Last Cell Count");
                        //}
                        //else
                        //{
                        //        System.out.println("Not the same " + "\n" + "iExcelLastCell: " + iExcelLastCell + "\n" + "lastcolumn: " + lastcolumn);
                        //}

                        int iRow = 0;

                        //cStmt = conn.prepareCall("{call Insert_PageTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}");

                        iExcelLastRow = worksheet.getLastRowNum();

                        for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                        {
                                indexRow = iExcelRow;
                                int paramindex = 2;
                                cStmt = conn.prepareCall(spcommand);
                                
                                cStmt.setNString(1, projcode);
                                cStmt.setString(paramindex, worksheetname);
                                conCatValue = paramindex + "|" + "STRING" + "-" + worksheetname + "\n";

                                for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                {
                                        indexCell = iExcelCell;
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
                    fis.close();

                    String errMessage = null;

                    errMessage = "Worksheet: " + worksheetname + ", " + "Row: " + indexRow + ", " + "Cell: " + indexCell + "\n" +
                        "Error Message: " + ex.getMessage() + 
                        "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                        "\n" + "Stack Trace: " + ex.getStackTrace() + 
                        "\n" + "Cause: " + ex.getCause();

                    System.out.println(errMessage);

                    JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                                   "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);
            }

    }
    
    public static void readUserMenuMap() throws Exception{
            JFrame jFrame = new JFrame();
            SQLObj sqlObj = new SQLObj();

            Connection conn = null;
            CallableStatement cStmt = null;
            String filename = "";
            String worksheetname = "";

            FileInputStream fis = null;
            String usercode = "";
            int intRow = 0;
            int intCell = 0;
            String sheetname = "";
            int paramindex = 0;
            
            try 
            {
                    conn = ewb.qa.tdd.SQLObj.ConnToDB();
                    cStmt = conn.prepareCall("{call Search_UserMenuMap_DistinctUserCode()}");
                   
                    ResultSet rs = cStmt.executeQuery();
                    int iRecordCount = rs.getRow();
                    String projcode = MainGUI.getProjectCode();
                    boolean flagNext = false;
                    
                    while(rs.next())
                    {
                            usercode = rs.getNString("USERCODE");
                            //String maincode = rs.getNString("MAINCODE");
                            //String subcode = rs.getNString("SUBCODE");
                            //String childcode = rs.getNString("CHILDCODE");
                            String spcommand = "{call Insert_MenuTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                            String qcommand = "";
                            String sequenceflag = "";
                            int lastcolumn = 0;

                            //if(usercode.equals("CBGUser0001")){
                            //***********************************************************************************************************************************************************************************
                            switch(usercode)
                            {
                            
                            case "User0001":
                                    //UniversalTellerMenu
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\UniversalTellerMenu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;

                            case "User0002":
                                    //ServiceManagerMenu
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\ServiceManagerMenu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;
                                    
                            case "LendingUser0001":
                                    //CBG Maker
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0001 - CBG Maker Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;
                                    
                            case "LendingUser0002":
                                    //CBG Maker
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0002 - Approver 348 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;

                            case "LendingUser0003":
                                    //CBG Maker
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0003 - Approver 310 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;                                    

                            case "LendingUser0004":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0004 - Approver 367 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;            
                                    
                            case "LendingUser0005":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0005 - Retail Credit Manager Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;         

                            case "LendingUser0006":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0006 - Approver 368 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    flagNext = false;
                                    break;                                      
                                
                            case "LendingUser0007":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0007 - Corporate Lending Supervisor.xlsx";
                                    System.out.println("User: " + usercode);
                                    flagNext = false;
                                    break;                                        
                                    
                            default:
                                flagNext = true;
                                //NextRecord;
                            }

                            if(flagNext == false){
                                String conCatValue = "";
                                String cellStringValue = "";
                                //FileInputStream fis = new FileInputStream(filename);
                                fis = new FileInputStream(filename);
                                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                                //worksheetname = maincode;
                                int intSheetMax = workbook.getNumberOfSheets() - 1;

                                //int sheetCtr = 0;
                                //XSSFSheet worksheet = workbook.getSheet(worksheetnamnt <= intSheetMax; sheetCount++){
                                        //String sheetname = workboe);
                                        
                                //blankrow:
                                for(int sheetCount = 1; sheetCount <= intSheetMax; sheetCount++){
                                        //String sheetname = workbook.getSheetAt(intSheetMax)
                                        XSSFSheet worksheet = workbook.getSheetAt(sheetCount);
                                        sheetname = worksheet.getSheetName();

                                        if(!sheetname.equals(usercode)){
                                                //******************************************************************************************************************
                                                if(sheetname.equals("CBGGChild00002_00001_00002")){
                                                    System.out.print("CBGGChild00002_00001_00002");
                                                }
                                                
                                                if(worksheet != null){
                                                        try
                                                        {
                                                                int iExcelLastRow = worksheet.getLastRowNum();
                                                                int iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;

                                                                for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                                                                {
                                                                        int iRow = 0;
                                                                        paramindex = 3;
                                                                        CallableStatement ins_cStmt = null;
                                                                        XSSFCell cell = worksheet.getRow(iExcelRow).getCell(0);
                                                                        if(cell == null ){
                                                                            break;
                                                                        }
                                                                        /*
                                                                        @proj_code NVARCHAR(20),
                                                                        @user_code NVARCHAR(20),
                                                                        @menu_code NVARCHAR(20),
                                                                        @childmenu_code NVARCHAR(20),
                                                                        @page_module NVARCHAR(50),
                                                                        @page_field NVARCHAR(50),
                                                                        @element_id NVARCHAR(200),
                                                                        @element_xpath NVARCHAR(200),
                                                                        @link_value NVARCHAR(50),
                                                                        @page_url NVARCHAR(200),
                                                                        @page_title NVARCHAR(100),
                                                                        @element_type NVARCHAR(50),
                                                                        @element_value NVARCHAR(100),
                                                                        @element_action NVARCHAR(50),
                                                                        @menu_reference NVARCHAR(50),
                                                                        @worksheet_reference NVARCHAR(50),
                                                                        @field_reference NVARCHAR(50),
                                                                        @element_message NVARCHAR(500),
                                                                        @element_result NVARCHAR(100),
                                                                        @menu_lvl INT,
                                                                        @sequence INT                                                                    
                                                                        */

                                                                        ins_cStmt = conn.prepareCall(spcommand);

                                                                        ins_cStmt.setNString(1, projcode);
                                                                        conCatValue = 1 + "|" + "STRING" + "-" + projcode + "\n";

                                                                        ins_cStmt.setString(2, usercode);
                                                                       conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + usercode + "\n";

                                                                        ins_cStmt.setString(3, sheetname);
                                                                        conCatValue = conCatValue + 3 + "|" + "STRING" + "-" + sheetname + "\n";

                                                                        for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                                                        {
                                                                                intRow = iExcelRow;
                                                                                intCell = iExcelCell;

                                                                                cellStringValue = "";
                                                                                int cellIntValue = 0;
                                                                                paramindex++;

                                                                                XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
//*************************************************************************************************************                                                                                        
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
                                                                                                        cellStringValue = NumberToTextConverter.toText(cellIntValue);
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
//*************************************************************************************************************                                                                                                                                                                                
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
                                                                String errMessage = null;

                                                                errMessage = "User Code: " + usercode + 
                                                                    "\n" + "Sheet Name: " + sheetname +
                                                                    "\n" + "Param Index : " + paramindex + ",  Row :" + intRow + ", Cell : " + intCell + ", Value : " + cellStringValue +
                                                                    "\n" + "Error Message: " + ex.getMessage() + 
                                                                    "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                                                                    "\n" + "Stack Trace: " + ex.getStackTrace() + 
                                                                    "\n" + "Cause: " + ex.getCause();

                                                                System.out.println(errMessage + "\n" + conCatValue);

                                                                JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                                                                        "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);   
                                                        }	
                                                }
                                                //*********************************************************************************************************************
                                        }
                                }                                
                                
                            }

                            /*
                            NextRecord:
                            if(flagNext == true){
                                //next record
                            }
                            */
                    }
                    cStmt.close();
                    conn.commit();
                    conn.close();

            } 
            catch (Exception ex) {
                    // TODO: handle exception
                    fis.close();
                    String errMessage = null;

                    errMessage = "User Code: " + usercode + 
                        "\n" + "Sheet Name: " + sheetname +
                        "\n" + "Param Index : " + paramindex + ",  Row :" + intRow + ", Cell: " + intCell +"Error Message: " + ex.getMessage() + 
                       "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                       "\n" + "Stack Trace: " + ex.getStackTrace() + 
                       "\n" + "Cause: " + ex.getCause();

                    System.out.println(errMessage);

                    JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                            "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);                    
            }
    }    

    public static boolean readUserMenuMap_ByUserCode(String givenUserCode, String givenFilePath) throws Exception{
            JFrame jFrame = new JFrame();
            SQLObj sqlObj = new SQLObj();

            Connection conn = null;
            CallableStatement cStmt = null;
            String filename = "";
            String worksheetname = "";

            FileInputStream fis = null;
            String usercode = "";
            int intRow = 0;
            int intCell = 0;
            String sheetname = "";
            int paramindex = 0;
            boolean result = false;
            
            try 
            {
                    result = true;
                    conn = ewb.qa.tdd.SQLObj.ConnToDB();
                    //cStmt = conn.prepareCall("{call Search_UserMenuMap_DistinctUserCode()}");
                    cStmt = conn.prepareCall("{call Search_UserMenuMap_DistinctUserCode_ByUserCode(?)}");
                    cStmt.setNString(1, givenUserCode);
                    
                    ResultSet rs = cStmt.executeQuery();
                    int iRecordCount = rs.getRow();
                    String projcode = MainGUI.getProjectCode();
                    boolean flagNext = false;
                    
                    while(rs.next())
                    {
                            usercode = rs.getNString("USERCODE");
                            //String maincode = rs.getNString("MAINCODE");
                            //String subcode = rs.getNString("SUBCODE");
                            //String childcode = rs.getNString("CHILDCODE");
                            String spcommand = "{call Insert_MenuTestscript(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}";
                            String qcommand = "";
                            String sequenceflag = "";
                            int lastcolumn = 0;

                            //if(usercode.equals("CBGUser0001")){
                            //***********************************************************************************************************************************************************************************
                            /*
                            switch(usercode)
                            {
                            
                            case "User0001":
                                    //UniversalTellerMenu
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\UniversalTellerMenu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;

                            case "User0002":
                                    //ServiceManagerMenu
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\ServiceManagerMenu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;
                                    
                            case "LendingUser0001":
                                    //CBG Maker
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0001 - CBG Maker Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;
                                    
                            case "LendingUser0002":
                                    //CBG Maker
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0002 - Approver 348 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;

                            case "LendingUser0003":
                                    //CBG Maker
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0003 - Approver 310 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;                                    

                            case "LendingUser0004":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0004 - Approver 367 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;            
                                    
                            case "LendingUser0005":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0005 - Retail Credit Manager Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    break;         

                            case "LendingUser0006":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0006 - Approver 368 Menu.xlsx";
                                    System.out.println("User: " + usercode);
                                    flagNext = false;
                                    break;                                      
                                
                            case "LendingUser0007":
                                    //CBG Authoriser
                                    filename = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver_Demo\\templates\\LendingUser0007 - Corporate Lending Supervisor.xlsx";
                                    System.out.println("User: " + usercode);
                                    flagNext = false;
                                    break;                                        
                                    
                            default:
                                flagNext = true;
                                //NextRecord;
                            }
                            */
                            
                            if(flagNext == false){
                                String conCatValue = "";
                                String cellStringValue = "";
                                //FileInputStream fis = new FileInputStream(filename);
                                //fis = new FileInputStream(filename);
                                
                                fis = new FileInputStream(givenFilePath);
                                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                                //worksheetname = maincode;
                                int intSheetMax = workbook.getNumberOfSheets() - 1;

                                //int sheetCtr = 0;
                                //XSSFSheet worksheet = workbook.getSheet(worksheetnamnt <= intSheetMax; sheetCount++){
                                        //String sheetname = workboe);
                                        
                                //blankrow:
                                for(int sheetCount = 1; sheetCount <= intSheetMax; sheetCount++){
                                        //String sheetname = workbook.getSheetAt(intSheetMax)
                                        XSSFSheet worksheet = workbook.getSheetAt(sheetCount);
                                        sheetname = worksheet.getSheetName();

                                        if(!sheetname.equals(usercode)){
                                                //******************************************************************************************************************
                                                //if(sheetname.equals("CBGGChild00002_00001_00002")){
                                                //    System.out.print("CBGGChild00002_00001_00002");
                                                //}
                                                
                                                if(worksheet != null){
                                                        try
                                                        {
                                                                int iExcelLastRow = worksheet.getLastRowNum();
                                                                int iExcelLastCell = worksheet.getRow(0).getLastCellNum() - 1;

                                                                for(int iExcelRow=1; iExcelRow<=iExcelLastRow; iExcelRow++)
                                                                {
                                                                        int iRow = 0;
                                                                        paramindex = 3;
                                                                        CallableStatement ins_cStmt = null;
                                                                        XSSFCell cell = worksheet.getRow(iExcelRow).getCell(0);
                                                                        if(cell == null ){
                                                                            break;
                                                                        }
                                                                        /*
                                                                        @proj_code NVARCHAR(20),
                                                                        @user_code NVARCHAR(20),
                                                                        @menu_code NVARCHAR(20),
                                                                        @childmenu_code NVARCHAR(20),
                                                                        @page_module NVARCHAR(50),
                                                                        @page_field NVARCHAR(50),
                                                                        @element_id NVARCHAR(200),
                                                                        @element_xpath NVARCHAR(200),
                                                                        @link_value NVARCHAR(50),
                                                                        @page_url NVARCHAR(200),
                                                                        @page_title NVARCHAR(100),
                                                                        @element_type NVARCHAR(50),
                                                                        @element_value NVARCHAR(100),
                                                                        @element_action NVARCHAR(50),
                                                                        @menu_reference NVARCHAR(50),
                                                                        @worksheet_reference NVARCHAR(50),
                                                                        @field_reference NVARCHAR(50),
                                                                        @element_message NVARCHAR(500),
                                                                        @element_result NVARCHAR(100),
                                                                        @menu_lvl INT,
                                                                        @sequence INT                                                                    
                                                                        */

                                                                        ins_cStmt = conn.prepareCall(spcommand);

                                                                        ins_cStmt.setNString(1, projcode);
                                                                        conCatValue = 1 + "|" + "STRING" + "-" + projcode + "\n";

                                                                        ins_cStmt.setString(2, usercode);
                                                                       conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + usercode + "\n";

                                                                        ins_cStmt.setString(3, sheetname);
                                                                        conCatValue = conCatValue + 3 + "|" + "STRING" + "-" + sheetname + "\n";

                                                                        for(int iExcelCell=0; iExcelCell<=iExcelLastCell; iExcelCell++)
                                                                        {
                                                                                intRow = iExcelRow;
                                                                                intCell = iExcelCell;

                                                                                cellStringValue = "";
                                                                                int cellIntValue = 0;
                                                                                paramindex++;

                                                                                XSSFCell cellObj = worksheet.getRow(iExcelRow).getCell(iExcelCell);
//*************************************************************************************************************                                                                                        
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
                                                                                                        cellStringValue = NumberToTextConverter.toText(cellIntValue);
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
//*************************************************************************************************************                                                                                                                                                                                
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
                                                                String errMessage = null;

                                                                errMessage = "User Code: " + usercode + 
                                                                    "\n" + "Sheet Name: " + sheetname +
                                                                    "\n" + "Param Index : " + paramindex + ",  Row :" + intRow + ", Cell : " + intCell + ", Value : " + cellStringValue +
                                                                    "\n" + "Error Message: " + ex.getMessage() + 
                                                                    "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                                                                    "\n" + "Stack Trace: " + ex.getStackTrace() + 
                                                                    "\n" + "Cause: " + ex.getCause();

                                                                System.out.println(errMessage + "\n" + conCatValue);

                                                                JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                                                                        "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);   
                                                        }	
                                                }
                                                //*********************************************************************************************************************
                                        }
                                }                                
                                
                            }
                    }
                    cStmt.close();
                    conn.commit();
                    conn.close();
                    return result;
            } 
            catch (Exception ex) {
                    // TODO: handle exception

                    String errMessage = null;
                    result = false;
                    
                    errMessage = "User Code: " + usercode + 
                        "\n" + "Sheet Name: " + sheetname +
                        "\n" + "Param Index : " + paramindex + ",  Row :" + intRow + ", Cell: " + intCell +"Error Message: " + ex.getMessage() + 
                       "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                       "\n" + "Stack Trace: " + ex.getStackTrace() + 
                       "\n" + "Cause: " + ex.getCause();

                    System.out.println(errMessage);

                    JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                            "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);   
                    
                    fis.close();
                    return result;
                    
            }
    }    
    
    public void setTestData(String filename, String loginid){
        JFrame jFrame = new JFrame();
        try{
                FileInputStream fis = new FileInputStream(filename);
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                XSSFSheet worksheet = null;
                
                int worksheetcount = workbook.getNumberOfSheets() - 1;
                
                /*
                	@proj_code NVARCHAR(20),
	@testcaseid NVARCHAR(12),
	@function_code  NVARCHAR(20),
	@field_id NVARCHAR(50),
	@field_value NVARCHAR(100),
	@iteration INT
                */
                
                Connection conn = ewb.qa.tdd.SQLObj.ConnToDB();
                
                for(int isheet = 0; isheet <= worksheetcount; isheet++){
                        worksheet = workbook.getSheetAt(isheet);
                        
                        String worksheetname = worksheet.getSheetName();
                        XSSFCell cellObj = null;
                        int LastRow = worksheet.getLastRowNum();
                        int LastCell = worksheet.getRow(3).getLastCellNum() - 1;
                        int Iteration = 1;

                        CallableStatement cStmt = null;
                        
                        String ProjCode = MainGUI.getProjectCode();
                        String TestCaseId = worksheet.getRow(1).getCell(1).getStringCellValue();
                        
                        for(int iCell = 1; iCell <= LastCell; iCell++){
                                int cellIntValue = 0;
                                String cellStringValue = "";
                                String conCatValue = "";
                                int paramindex = 4;

                                for(int iRow = 3; iRow <= LastRow; iRow++){
                                        cStmt = conn.prepareCall("{call Insert_Testdata_ByLoginId(?,?,?,?,?,?,?)}");
                                        
                                        cStmt.setNString(1, ProjCode);
                                        conCatValue = 1 + "|" + ProjCode + "\n";

                                        cStmt.setNString(2, TestCaseId);
                                        conCatValue = conCatValue + 2 + "|" + TestCaseId + "\n";

                                        cStmt.setNString(3, worksheetname);
                                        conCatValue = conCatValue + 3 + "|" + worksheetname + "\n";

                                        String fieldid = "";                                    
                                        fieldid = worksheet.getRow(iRow).getCell(0).getStringCellValue();
                                        cStmt.setNString(4, fieldid);
                                        conCatValue = conCatValue + 4 + "|" + fieldid + "\n";

                                        cellObj = worksheet.getRow(iRow).getCell(iCell);
                                        if(cellObj == null){
                                                cellStringValue = "null";
                                        }
                                        else{
                                            switch(cellObj.getCellType()){
                                                case STRING:
                                                    cellStringValue = cellObj.getStringCellValue();                                                    
                                                    break;
                                                    
                                                case NUMERIC:
                                                    cellStringValue = NumberToTextConverter.toText(cellObj.getNumericCellValue());
                                                    break;
                                                    
                                                case BLANK:
                                                    cellStringValue = "null";
                                                    break;
                                                case ERROR:
                                                    cellStringValue = "null";
                                                    break;                                                    
                                                    
                                            }

                                        }
                                        
                                        cStmt.setNString(5, cellStringValue);
                                        conCatValue = conCatValue + 5 + "|" + cellStringValue + "\n";

                                        cStmt.setInt(6, Iteration);
                                        conCatValue = conCatValue + 6 + "|" + Iteration + "\n";

                                        cStmt.setNString(7, loginid);
                                        conCatValue = conCatValue + 7 + "|" + loginid + "\n";
                                        
                                        cStmt.execute();
                                        conn.commit();
                                        cStmt.close();
                                        System.out.println(conCatValue);

                                }
                                Iteration++;
                        }

                }
                conn.close();   
        }
        
        catch(Exception ex){
                String errMessage = null;

                errMessage = "Error Message: " + ex.getMessage() + 
                   "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                   "\n" + "Stack Trace: " + ex.getStackTrace() + 
                   "\n" + "Cause: " + ex.getCause();

                System.out.println(errMessage);

                JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                        "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);    
        }
        
    }
    
    //MapElementType(String Module, String Field, String Title, String Elementtype, String Elementid, String Elementxpath, String Elementvalue, String Action, String ReferenceField, boolean TakeScreenshot, String testcaseid)
    public static boolean MapElementType(String Module, String Field, String Elementtype, String Elementid, String Elementxpath, String Elementvalue, String Action, 
            String ReferenceField, boolean TakeScreenshot, String testcaseid, String usercode, String funccode, String pagemodule, int testcycle, int testrun, String projcode, String loginid, String LinkText, String PageTitle){
            JFrame jFrame = new JFrame();
            seleniumObj = new SeleniumObj();
            String resultMsg = "Test Case Id :" + testcaseid + "| "+ "Module :" + Module + "| " + "Field :" + Field + "| " + "Element Type :" + Elementtype + "| " + 
                    "Element Id:" + Elementid + "| " + "Element Xpath :" + Elementxpath + "| " + "Element Value :" + Elementvalue + "| " + "Action :" + Action + "|";
            
            boolean MapElementResult = true;
            String errMessage = null;
            
            DateFormat dateformat = new SimpleDateFormat("MM/dd/yyy HH:mm:ss");
            Date currentdate = new Date();
            String systemdate = dateformat.format(currentdate);  
            //dateformat.format(currentdate)
            Date convertdate = null;
            String TagType = "";
            String TagName = "";
            String ValueRef = "null";
            
            String varElem = "";
            String varElemValue = "";
            String varElemType = Elementtype;
            String varElemAction = Action;
            
//            try
//            {
                    try{
                        convertdate = new SimpleDateFormat("MM/dd/yyy HH:mm:ss").parse(systemdate);
                    }
                    catch(Exception ex){
                        errMessage = "Object convertdate - failed " + "\n" +
                            "Error Message :" + ex.getMessage() + "\n" +
                            "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                            "Stack Trace: " + ex.getStackTrace() + "\n" +
                            "Cause: " + ex.getCause();

                        MapElementResult = false;
                        SeleniumObj.setEventErrorMsg(errMessage);                        
                    }

                    switch(Elementtype)
                    {
                    case "Input Box":
                            TagName = "input";
                            TagType = "text";
                            String currentWindowTitle = "";
                            
                            try{
                                if(!Elementvalue.equals("null") && !ReferenceField.equals("null"))
                                {
                                        StoreToTemp(Elementvalue, ReferenceField, LinkText);
                                }
                                else if(Elementvalue.equals("null") && !ReferenceField.equals("null"))
                                {
                                        Elementvalue = ReadToTemp(ReferenceField);
                                }
                            
                                if(!Elementid.equals("null"))
                                {
                                    varElem = Elementid;
                                    varElemValue = Elementvalue;
                                    MapElementResult = seleniumObj.SendKeys_EventById(Elementid, Elementvalue, ReferenceField);
                                }
                                else if(!Elementxpath.equals("null"))
                                {
                                    varElem = Elementxpath;
                                    varElemValue = Elementvalue;
                                    MapElementResult = seleniumObj.SendKeys_EventByXpath(Elementxpath, Elementvalue, ReferenceField);
                                }                                
                            }
                            catch(Exception ex){
                                errMessage = "Element :" + varElem + "\n" + 
                                    "Element Type :" + varElemType + ", " +
                                    "Element Value :" + varElemValue + ", " +
                                    "Element Action :" + varElemAction + ", " + "\n" +
                                    "Error Message :" + ex.getMessage() + "\n" +
                                    "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                    "Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Cause: " + ex.getCause();

                                MapElementResult = false;
                                SeleniumObj.setEventErrorMsg(errMessage);
                            }
                            finally{
                                if(MapElementResult == false){
                                    try{
                                        currentWindowTitle = SeleniumObj.GetTitle_NewTab();
                                        if(currentWindowTitle.equals(PageTitle)){
                                            SeleniumObj.getCurrentWindow(PageTitle);                                            
                                        }

                                        if(!Elementid.equals("null"))
                                        {
                                            MapElementResult = seleniumObj.SendKeys_EventById(Elementid, Elementvalue, ReferenceField);
                                        }
                                        else if(!Elementxpath.equals("null"))
                                        {
                                            MapElementResult = seleniumObj.SendKeys_EventByXpath(Elementxpath, Elementvalue, ReferenceField);
                                        }                                                      
                                    }
                                    catch(Exception ex1){
                                        errMessage = "Element :" + varElem + "\n" + 
                                            "Element Type :" + varElemType + ", " +
                                            "Element Value :" + varElemValue + ", " +
                                            "Element Action :" + varElemAction + ", " + "\n" +
                                            "Error Message :" + ex1.getMessage() + "\n" +
                                            "Error Localize Message: " + ex1.getLocalizedMessage() + "\n" +
                                            "Stack Trace: " + ex1.getStackTrace() + "\n" +
                                            "Cause: " + ex1.getCause();

                                        MapElementResult = false;
                                        SeleniumObj.setEventErrorMsg(errMessage);                                        
                                    }
                                   
                                }
                            }

                            break;

                    case "Button":
                            TagName = "img";
                            TagType = "";
                            
                            try{
                                if(!Elementid.equals("null"))
                                {
                                    varElem = Elementid;
                                    varElemValue = Elementvalue;
                                    MapElementResult = seleniumObj.Click_EventById(Elementid);
                                        
                                }
                                else if(!Elementxpath.equals("null"))
                                {
                                    varElem = Elementid;
                                    varElemValue = Elementvalue;
                                    MapElementResult = seleniumObj.Click_EventByXpath(Elementxpath);
                                }                                
                            }
                            catch(Exception ex){
                                errMessage = "Element :" + varElem + "\n" + 
                                    "Element Type :" + varElemType + ", " +
                                    "Element Value :" + varElemValue + ", " +
                                    "Element Action :" + varElemAction + ", " + "\n" +
                                    "Error Message :" + ex.getMessage() + "\n" +
                                    "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                    "Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Cause: " + ex.getCause();
                                MapElementResult = false;
                                SeleniumObj.setEventErrorMsg(errMessage);
                            }
                            finally{
                                if(MapElementResult == false){
                                    try{
                                        currentWindowTitle = SeleniumObj.GetTitle_NewTab();
//                                        if(currentWindowTitle.equals(PageTitle)){
//                                            SeleniumObj.getCurrentWindow(PageTitle);                                            
//                                        }
                                        SeleniumObj.getCurrentWindow(currentWindowTitle);
                                        
                                        if(!Elementid.equals("null"))
                                        {
                                            MapElementResult = seleniumObj.Click_EventById(Elementid);
                                        }
                                        else if(!Elementxpath.equals("null"))
                                        {
                                            MapElementResult = seleniumObj.Click_EventByXpath(Elementxpath);
                                        }                                                      
                                    }
                                    catch(Exception ex1){
                                        errMessage = "Element :" + varElem + "\n" + 
                                            "Element Type :" + varElemType + ", " +
                                            "Element Value :" + varElemValue + ", " +
                                            "Element Action :" + varElemAction + ", " + "\n" +
                                            "Error Message :" + ex1.getMessage() + "\n" +
                                            "Error Localize Message: " + ex1.getLocalizedMessage() + "\n" +
                                            "Stack Trace: " + ex1.getStackTrace() + "\n" +
                                            "Cause: " + ex1.getCause();

                                        MapElementResult = false;
                                        SeleniumObj.setEventErrorMsg(errMessage);          
                                    }
                                }            
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
                            int iframeNum = 0;
                            //Integer.parseInt(Elementvalue);
                            try{
                                if(!Elementid.equals("null"))
                                {
                                    switch(Elementvalue){
                                        case "1":
                                            iframeNum = 1;
                                            break;

                                        case "2":
                                            iframeNum = 2;
                                            break;

                                        case "3":
                                            iframeNum = 3;
                                            break;
                                    }                                
                                    MapElementResult = seleniumObj.Select_EventById(Elementid, iframeNum);
                                }
                                else if(!Elementxpath.equals("null"))
                                {
                                    switch(Elementvalue){
                                        case "1":
                                            iframeNum = 1;
                                            break;

                                        case "2":
                                            iframeNum = 2;
                                            break;

                                        case "3":
                                            iframeNum = 3;
                                            break;
                                    }
                                    MapElementResult = seleniumObj.Select_EventByXpath(Elementxpath, iframeNum);
                                }                                
                            }
                            catch(Exception ex){
                                errMessage = "Element :" + varElem + "\n" + 
                                    "Element Type :" + varElemType + ", " +
                                    "Element Value :" + varElemValue + ", " +
                                    "Element Action :" + varElemAction + ", " + "\n" +
                                    "Error Message :" + ex.getMessage() + "\n" +
                                    "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                    "Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Cause: " + ex.getCause();
                                MapElementResult = false;
                                SeleniumObj.setEventErrorMsg(errMessage);     
                            }

                            break;

                    case "Browse":
                        try{
                            String strTitleNewTab;
                            strTitleNewTab = seleniumObj.Title_NewTab();
                            SeleniumObj.getCurrentWindow(strTitleNewTab);
                            System.out.println(strTitleNewTab);                            
                        }
                        catch(Exception ex){
                                errMessage = "Element :" + varElem + "\n" + 
                                    "Element Type :" + varElemType + ", " +
                                    "Element Value :" + varElemValue + ", " +
                                    "Element Action :" + varElemAction + ", " + "\n" +
                                    "Error Message :" + ex.getMessage() + "\n" +
                                    "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                    "Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Cause: " + ex.getCause();
                                MapElementResult = false;
                                SeleniumObj.setEventErrorMsg(errMessage);                                
                        }

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

                    case "NewWindow":
                            switch(Action){
                                case "Select":
                                    MapElementResult = seleniumObj.NewWindow_Select(PageTitle);
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
                        try{
                            if(!Elementid.equals("null"))
                            {
                                varElem = Elementid;
                                varElemValue = Elementvalue;
                                MapElementResult = seleniumObj.DropdownBox_ById(Elementid, Elementvalue, ReferenceField);
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                varElem = Elementxpath;
                                varElemValue = Elementvalue;
                                MapElementResult = seleniumObj.DropdownBox_ByXpath(Elementxpath, Elementvalue, ReferenceField);
                            }                            
                        }
                        catch(Exception ex){
                            errMessage = "Element :" + varElem + "\n" + 
                                "Element Type :" + varElemType + ", " +
                                "Element Value :" + varElemValue + ", " +
                                "Element Action :" + varElemAction + ", " + "\n" +
                                "Error Message :" + ex.getMessage() + "\n" +
                                "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                "Stack Trace: " + ex.getStackTrace() + "\n" +
                                "Cause: " + ex.getCause();
                            MapElementResult = false;
                            SeleniumObj.setEventErrorMsg(errMessage);                                  
                        }
                                
                        break;

                    case "Text Message":
                        try{
                            if(!Elementid.equals("null"))
                            {
                                if(Action.equals("GetText")){
                                    if(!ReferenceField.equals("null")){
                                        varElem = Elementid;
                                        MapElementResult = SeleniumObj.GetInputText_EventById(Elementid, ReferenceField);
                                    }
                                }
                            }
                            else if(!Elementxpath.equals("null"))
                            {
                                if(Action.equals("GetText")){
                                    if(!ReferenceField.equals("null")){
                                        varElem = Elementxpath;
                                        MapElementResult = SeleniumObj.GetInputText_EventByXpath(Elementxpath, ReferenceField);
                                    }
                                }
                            }                            
                        }
                        catch(Exception ex){
                                errMessage = "Element :" + varElem + "\n" + 
                                    "Element Type :" + varElemType + ", " +
                                    "Element Value :" + varElemValue + ", " +
                                    "Element Action :" + varElemAction + ", " + "\n" +
                                    "Error Message :" + ex.getMessage() + "\n" +
                                    "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                    "Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Cause: " + ex.getCause();
                                MapElementResult = false;
                                SeleniumObj.setEventErrorMsg(errMessage);                            
                        }

                        break;
                            
                    case "Radio":
                            TagName = "input";
                            TagType = "radio";
                            boolean elemFound = false;
                            
                            try{
                                if(!Elementid.equals("null"))
                                {
                                        varElem = Elementid;
                                        MapElementResult = seleniumObj.Click_EventById(Elementid);
                                }
                                else if(!Elementxpath.equals("null"))
                                {
                                        varElem = Elementxpath;
                                        MapElementResult = seleniumObj.Click_EventByXpath(Elementxpath);
                                }                                
                            }
                            catch(Exception ex){
                                errMessage = "Element :" + varElem + "\n" + 
                                    "Element Type :" + varElemType + ", " +
                                    "Element Value :" + varElemValue + ", " +
                                    "Element Action :" + varElemAction + ", " + "\n" +
                                    "Error Message :" + ex.getMessage() + "\n" +
                                    "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                    "Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Cause: " + ex.getCause();
                                MapElementResult = false;
                                SeleniumObj.setEventErrorMsg(errMessage);                                        
                            }

                            break;                 
                            
                    case "SelectWindow":
                        try{
                            seleniumObj.Select_Window();                                
                        }
                        catch(Exception ex){
                            errMessage = "Element :" + varElem + "\n" + 
                                "Element Type :" + varElemType + ", " +
                                "Element Value :" + varElemValue + ", " +
                                "Element Action :" + varElemAction + ", " + "\n" +
                                "Error Message :" + ex.getMessage() + "\n" +
                                "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
                                "Stack Trace: " + ex.getStackTrace() + "\n" +
                                "Cause: " + ex.getCause();
                            MapElementResult = false;
                            SeleniumObj.setEventErrorMsg(errMessage);                               
                        }

                        break;
                            
                    default:
                        MapElementResult = false;
                        errMessage = "Element Type is invalid" + "\n" +
                                "Field :" + Field + "\n" + "Element ID :" + Elementid + "\n" + "Element Xpath :" + Elementxpath + "\n" + "Element Type :" + Elementtype + "\n" + 
                                "Element Action :" + Action;
                        SeleniumObj.setEventErrorMsg(errMessage);
                    }			

                    if(TakeScreenshot == true)
                    {
                        try{
                            TakeScreenShot(seleniumObj.getWebDriver(), testcaseid, Module, Field);                            
                        }
                        catch(IOException ex){
                            String errIOMessage = null;

                            errMessage = "Error Message: " + ex.getMessage() + 
                               "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                               "\n" + "Stack Trace: " + ex.getStackTrace() + 
                               "\n" + "Cause: " + ex.getCause();

                            System.out.println(errIOMessage);

                            JOptionPane.showMessageDialog(jFrame,  "Screenshot failed" + 
                                           "\n" + errMessage, "Test Case Execution", JOptionPane.ERROR_MESSAGE);                            
                        }

                    }
                    
                    if(MapElementResult == false){
                        errMessage = SeleniumObj.getEventErrorMsg();
                    }
                    else{
                        errMessage = "Successfully executed";                        
                    }
                    
                    if(!ReferenceField.equals("null")){
                        ValueRef = getValueRef();                        
                    }
                    else{
                        ValueRef = "null";
                    }

                    //InsertToLog(
                    //String ProjCode, String TestCaseId, String UserCode, String FuncCode, String PageModule, String FieldId, String ElementId, String ElementXpath,
                    //String ElementType, String ElementValue, String ElementAction, int TestCycle, int TestRun, String Result, String ErrMessage, Date DTime)
                    InsertToLog(projcode, testcaseid, usercode, funccode, pagemodule, Field, Elementid, Elementxpath, Elementtype,  Elementvalue, Action, testcycle, testrun, 
                            Boolean.toString(MapElementResult), errMessage, loginid, ValueRef);
//            }
//            catch(Exception ex)
//            {
////                    String varElem = "";
////                    String varElemType = Elementtype;
////                    String varElemValue = Elementvalue;
////                    String varElemAction = Action;
//                    
//                    if(!Elementid.equals("null")){
//                        varElem = Elementid;
//                    }
//                    else if(!Elementxpath.equals("null")){
//                        varElem = Elementxpath;
//                    }
//                    errMessage = "Element :" + varElem + "\n" + 
//                        "Element Type :" + varElemType + ", " +
//                        "Element Value :" + varElemValue + ", " +
//                        "Element Action :" + varElemAction + ", " + "\n" +
//                        "Error Message :" + ex.getMessage() + "\n" +
//                        "Error Localize Message: " + ex.getLocalizedMessage() + "\n" +
//                        "Stack Trace: " + ex.getStackTrace() + "\n" +
//                        "Cause: " + ex.getCause();
//
//                    System.out.println(errMessage);
//
//                    //JOptionPane.showMessageDialog(jFrame,  "Browser loading failed" + 
//                    //               "\n" + errMessage + "\n" + resultMsg, "Test Case Execution", JOptionPane.ERROR_MESSAGE);
//                    MapElementResult = false;
//                    
//                    //InsertToLog(
//                    //String ProjCode, String TestCaseId, String UserCode, String FuncCode, String PageModule, String FieldId, String ElementId, String ElementXpath,
//                    //String ElementType, String ElementValue, String ElementAction, int TestCycle, int TestRun, String Result, String ErrMessage, Date DTime)
//                    InsertToLog(projcode, testcaseid, usercode, funccode, pagemodule, Field, Elementid, Elementxpath, Elementtype,  Elementvalue, Action, testcycle, testrun, Boolean.toString(MapElementResult), errMessage, loginid, ValueRef);
//            }
            return MapElementResult;
    }    
    
    public static String ReadToTemp(String fieldreference) throws Exception{
            ewb.qa.tdd.Properties propObj = new ewb.qa.tdd.Properties();
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
            setValueRef(value);
            return value;
    }    
    
    public static void StoreToTemp(String value, String fieldreference, String Type) throws Exception{
            ewb.qa.tdd.Properties propObj = new ewb.qa.tdd.Properties();
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
                    setValueRef(givenValue);
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
                                            else if(scanValue.equals(fieldreference)){
                                                if(Type.equals("Single")){
                                                    scanValue = value;
                                                    editFlag = true;
                                                    break;                                                    
                                                }
                                                else if(Type.equals("Multiple")){
                                                    scanValue = scanValue + "*" + value;
                                                    editFlag = true;
                                                    break;   
                                                }

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

    public static void TakeScreenShot(WebDriver webdriver, String testcaseid, String module, String field) throws IOException{
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

    }    
    
   public static void InsertToLog(String ProjCode, String TestCaseId, String UserCode, String FuncCode, String PageModule, String FieldId, String ElementId, String ElementXpath,
           String ElementType, String ElementValue, String ElementAction, int TestCycle, int TestRun, String Result, String ErrMessage, String LoginId, String ValueRef){
       JFrame jFrame = new JFrame();
       int executeflag = 0;
       //SimpleDateFormat dtformatter = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
       //Date currentdate  = new Date();
       
       try{
           //currentdate = dtformatter.parse(DTime);
           
           Connection conn = ewb.qa.tdd.SQLObj.ConnToDB();
           CallableStatement cStmt = conn.prepareCall("{call Insert_ExecutionLog(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)}");

           /*
            1-@proj_code NVARCHAR(20),
            2-@testcaseid  NVARCHAR(12),
            3-@user_code NVARCHAR(20),
            4-@menu_code NVARCHAR(20),
            5-@page_module NVARCHAR(50),
            6-@page_field NVARCHAR(50),
            7-@element_id NVARCHAR(200),
            8-@element_xpath NVARCHAR(200),
            9-@element_type NVARCHAR(50),
            10-@element_value NVARCHAR(100),
            11-@element_action NVARCHAR(50),
            12-@test_cycle INT,
            13-@test_run INT,
            14-@element_result NVARCHAR(100),
            15-@err_message TEXT
            16-@loginid NVARCHAR(20),
            17-@value_ref NVARCHAR(50)
           */
           
           cStmt.setNString(1, ProjCode);
           cStmt.setNString(2, TestCaseId);
           cStmt.setNString(3, UserCode);
           cStmt.setNString(4, FuncCode);
           cStmt.setNString(5, PageModule);
           cStmt.setNString(6, FieldId);
           cStmt.setNString(7, ElementId);
           cStmt.setNString(8, ElementXpath);
           cStmt.setNString(9, ElementType);
           cStmt.setNString(10, ElementValue);
           cStmt.setNString(11, ElementAction);
           cStmt.setInt(12, TestCycle);
           cStmt.setInt(13, TestRun);
           cStmt.setNString(14, Result);
           cStmt.setNString(15, ErrMessage);
           cStmt.setNString(16, LoginId);
           cStmt.setNString(17, ValueRef);
           
           executeflag = cStmt.executeUpdate();
           
           if(executeflag < 0){
               System.out.println("Inserted Test Log");
           }
       }
       catch(Exception ex){
                String errMessage = null;
                errMessage = "Error Message: " + ex.getMessage() + 
                   "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                   "\n" + "Stack Trace: " + ex.getStackTrace() + 
                   "\n" + "Cause: " + ex.getCause();

                System.out.println(errMessage);

                JOptionPane.showMessageDialog(jFrame,  "Insert Execution Log failed" + 
                               "\n" + errMessage, "Test Case Execution", JOptionPane.ERROR_MESSAGE);     
       }
   }
    
   public static void UploadProductMenuMap(String filename, String worksheetname, String UserCode, String ProjCode, String TblName) throws IOException{
       JFrame jFrame = new JFrame();
       FileInputStream fis = null;
       int executeflag = 0;
       
       try{
                fis = new FileInputStream(filename);
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                XSSFSheet worksheet = workbook.getSheet(worksheetname);

                if(worksheet != null){
                        int MaxIndexRow = worksheet.getLastRowNum();
                        int MaxIndexCell = worksheet.getRow(0).getLastCellNum() - 1;
                        Connection conn = null;
                        
                        for(int indexRow = 1; indexRow <= MaxIndexRow; indexRow++){
                            conn = ewb.qa.tdd.SQLObj.ConnToDB();
                            CallableStatement cStmt = null;
                            int paramindex = 0;
                            String conCatValue = "";
                            String SubMenuCode = "null";
                            String ChildMenuCode = "null";
                            int iRow = 0;
                            int flag = 0;
                            
                            switch(TblName){
                                 case "CATEGORY_MENUMAP":
                                        cStmt = conn.prepareCall("{call Insert_CategoryMenuMap(?,?,?,?,?,?,?,?,?,?,?)}");

                                        /*
                                        @proj_code NVARCHAR(20),
                                        @user_code NVARCHAR(20),
                                        @category_id NVARCHAR(20),
                                        @sub_menucode NVARCHAR(20),
                                        @child_menucode NVARCHAR(20),
                                        @module NVARCHAR(50),
                                        @field_id NVARCHAR(50),
                                        @element_id NVARCHAR(200),
                                        @element_xpath NVARCHAR(200),
                                        @element_type NVARCHAR(50),
                                        @element_action NVARCHAR(50)                
                                        */

                                        cStmt.setNString(1, ProjCode);
                                        conCatValue = 1 + "|" + "STRING" + "-" + ProjCode + "\n";

                                        cStmt.setNString(2, UserCode);
                                        conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + UserCode + "\n";                                     
                                        paramindex = 2;
                                        flag = 1;
                                        
                                     break;
                                     
                                 case "GROUP_MENUMAP":
                                     cStmt = conn.prepareCall("{call Insert_GroupMenuMap(?,?,?,?,?,?,?,?,?,?)}");
                                     /*
                                    1-@proj_code NVARCHAR(20),
                                    2-@user_code NVARCHAR(20),
                                    3-@category_id NVARCHAR(50),
                                    4-@group_id NVARCHAR(50),
                                    5-@module NVARCHAR(50),
                                    6-@field_id NVARCHAR(50),
                                    7-@element_id NVARCHAR(200),
                                    8-@element_xpath NVARCHAR(200),
                                    9-@element_type NVARCHAR(50),
                                    10-@element_action NVARCHAR(50)                                 
                                     */
                                     
                                     cStmt.setNString(1, ProjCode);
                                     conCatValue = 1 + "|" + "STRING" + "-" + ProjCode + "\n";
                                     
                                     cStmt.setNString(2, UserCode);
                                     conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + UserCode + "\n"; 
                                     
                                     cStmt.setNString(3, worksheetname);
                                     conCatValue = conCatValue + 3 + "|" + "STRING" + "-" + worksheetname + "\n"; 
                                     
                                     paramindex = 3;
                                     flag = 2;
                                     
                                     break;
                                     
                                 case "PRODUCT_MENUMAP":
                                     cStmt = conn.prepareCall("{call Insert_ProductMenuMap(?,?,?,?,?,?,?,?,?,?)}");                                     
                                     /*
                                    @proj_code NVARCHAR(20),
                                    @user_code NVARCHAR(20),
                                    @group_id NVARCHAR(20),
                                    @grpmenu_code NVARCHAR(20),
                                    @module NVARCHAR(50),
                                    @field_id NVARCHAR(50),
                                    @element_id NVARCHAR(200),
                                    @element_xpath NVARCHAR(200),
                                    @element_type NVARCHAR(50),
                                    @element_action NVARCHAR(50)                                     
                                     */
                                     cStmt.setNString(1, ProjCode);
                                     conCatValue = 1 + "|" + "STRING" + "-" + ProjCode + "\n";
                                     
                                     cStmt.setNString(2, UserCode);
                                     conCatValue = conCatValue + 2 + "|" + "STRING" + "-" + UserCode + "\n"; 
                                     
                                     cStmt.setNString(3, worksheetname);
                                     conCatValue = conCatValue + 3 + "|" + "STRING" + "-" + worksheetname + "\n"; 
                                     
                                     paramindex = 3;                       
                                     flag = 3;
                                     break;
                             }

                             for(int IndexCell=0; IndexCell<=MaxIndexCell; IndexCell++)
                             {
                                 if(flag == 1){
                                     //CATEGORY_MENUMAP
                                    if(IndexCell == 0 || IndexCell == 3 || IndexCell == 4 || IndexCell == 5 || IndexCell == 6 || IndexCell == 10 || IndexCell == 12){
                                            paramindex++;
                                            XSSFCell cellObj = worksheet.getRow(indexRow).getCell(IndexCell);
                                            String cellStringValue = "";
                                            int cellIntValue = 0;

                                            if(cellObj == null)
                                            {
                                                    cellStringValue = cellObj.getStringCellValue();
                                                    cStmt.setString(paramindex, cellStringValue);
                                            }
                                            else{
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

                                    }
                                    else if(IndexCell == 1){
                                        paramindex++;
                                        cStmt.setNString(paramindex, SubMenuCode);

                                    }
                                    else if(IndexCell == 2){
                                        paramindex++;
                                        cStmt.setNString(paramindex, ChildMenuCode);

                                    }                                     

                                 }
                                 else if(flag == 2 || flag == 3){
                                     //GROUP_MENUMAP || PRODUCT_MENUMAP
                                    if(IndexCell == 0 || IndexCell == 1 || IndexCell == 2 || IndexCell == 3 || IndexCell == 4 || IndexCell == 8 || IndexCell == 10){
                                            paramindex++;
                                            XSSFCell cellObj = worksheet.getRow(indexRow).getCell(IndexCell);
                                            String cellStringValue = "";
                                            int cellIntValue = 0;

                                            if(cellObj == null)
                                            {
                                                    cellStringValue = cellObj.getStringCellValue();
                                                    cStmt.setString(paramindex, cellStringValue);
                                            }
                                            else{
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

                                    }
                                     
                                 }

                                 iRow = IndexCell;
                                 //paramindex++;    
                             }

                             System.out.println(conCatValue);
                             executeflag = cStmt.executeUpdate();
                             if(executeflag < 0){
                                 System.out.println("Table : " + TblName + " - Record insert successfully");
                             }
                             else{
                                 System.out.println("Table : " + TblName + " - Record insert failed");
                             }
                             conn.commit();
                             cStmt.close();
                        }
                        conn.close();
                }
                fis.close();
       }
       catch(Exception ex){
                String errMessage = null;

                errMessage = "Error Message: " + ex.getMessage() + 
                   "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                   "\n" + "Stack Trace: " + ex.getStackTrace() + 
                   "\n" + "Cause: " + ex.getCause();

                System.out.println(errMessage);

                JOptionPane.showMessageDialog(jFrame,  "Record Saving Failed " + 
                               "\n" + errMessage, "Maintenance - Project Location", JOptionPane.ERROR_MESSAGE);      

                if(executeflag == 0){
                    System.out.println("Table : " + TblName + " - Record insert failed");
                }                
                
                fis.close();
       }
   }
                
   private static String globalValueRef;
   public static void setValueRef(String givenValue){
       globalValueRef = givenValue;
   }
   
   public static String getValueRef(){
       return globalValueRef;
   }
   
}
