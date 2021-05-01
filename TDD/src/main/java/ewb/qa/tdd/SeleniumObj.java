/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ewb.qa.tdd;

import com.thoughtworks.selenium.Selenium;
import ewb.qa.tdd.GUI.ExcelObject;
import static java.awt.SystemColor.window;
import java.io.IOException;
import java.net.URL;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Set;
import java.util.concurrent.TimeUnit;
import java.util.List;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.WebDriver.Timeouts;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.WebDriver.*;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.Command;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.Response;
import org.openqa.selenium.WebDriver.Options;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.remote.SessionId;

public class SeleniumObj {
Properties properties;
        static WebDriver driver;
        ExcelObj excelObj;
        private static String strbrowsermessage;
        private static String strnewtabtitle;
        private static boolean boolSwitchto = false;
        private static String strtextmessage;
        private static String eventerrormsg;

        public static void Loginto(String BrowserName, String BaseUrl) throws Exception{
                        JFrame jFrame = new JFrame();
                        //DesiredCapabilities capabilities = new DesiredCapabilities();
                        //capabilities.setCapability(CapabilityType.ForSeleniumServer.ENSURING_CLEAN_SESSION, true);
                        //ChromeDriver driver = new ChromeDriver(capabilities);
                        
                        try{
                                switch(BrowserName)
                                {
                                        case "Google Chrome":
                                                String drvpath87 = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\chrome87\\chromedriver.exe";
                                                String drvpath79 = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\chrome79\\chromedriver.exe";
                                                String webdrv = "webdriver.chrome.driver";

                                                System.setProperty(webdrv, drvpath87);
                                                driver = new ChromeDriver();
                                                driver.manage().deleteAllCookies();
                                                break;

                                        case "IE":
                                                System.setProperty("webdriver.chrome.driver", "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\IEDriverServer.exe");
                                                driver = new InternetExplorerDriver();
                                                driver.manage().deleteAllCookies();
                                                break;
                                }
                                driver.get(BaseUrl);

                                //driver.manage().timeouts().implicitlyWait(20,TimeUnit.SECONDS);
                                //driver.manage().window().maximize();			                            
                        }
                        catch(Exception ex){
                                String errMessage = null;

                                errMessage = "Error Message: " + ex.getMessage() + 
                                   "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                                   "\n" + "Stack Trace: " + ex.getStackTrace() + 
                                   "\n" + "Cause: " + ex.getCause();

                                System.out.println(errMessage);

                                JOptionPane.showMessageDialog(jFrame,  "Saving Failed " + 
                                               "\n" + errMessage, "Maintenance", 0);
                        }

        }

        public void testconnection(String BaseUrl) throws WebDriverException {
                JFrame jFrame = new JFrame();
                try{   
                        //System.setProperty("webdriver.chrome.driver", "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\chrome87\\chromedriver.exe");
                        String drvpath87 = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\chrome87\\chromedriver.exe";
                        String drvpath79 = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\chrome79\\chromedriver.exe";
                        String drvpath89 = "D:\\QA\\Projects\\Test Automation\\Java_SeleniumWebDriver\\webbrowser_driver\\chrome89\\chromedriver.exe";
                        
                        String webdrv = "webdriver.chrome.driver";

                        System.setProperty(webdrv, drvpath89);
                        driver = new ChromeDriver();
                        
                        driver.manage().deleteAllCookies();
                        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                        driver.manage().window().maximize();  
                        driver.get(BaseUrl);
                        
                        SessionId sessionId = ((RemoteWebDriver)driver).getSessionId();
                        //sessionId.
                        
                        setWebDriver(driver);

                        //driver = new ChromeDriver();  
                       driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                        
                        //driver.get(BaseUrl);  

                        String URL= driver.getCurrentUrl();  
                        System.out.print(URL);  

                        String title = driver.getTitle();                  
                        System.out.println(title);  
                }
                catch(WebDriverException ex){
                        String errMessage = null;

                        errMessage = "Error Message: " + ex.getMessage() + 
                           "\n" + "Error Localize Message: " + ex.getLocalizedMessage() + 
                           "\n" + "Stack Trace: " + ex.getStackTrace() + 
                           "\n" + "Cause: " + ex.getCause();

                        System.out.println(errMessage);

                        JOptionPane.showMessageDialog(jFrame,  "Failed to Load browser " + 
                                       "\n" + errMessage, "Test Case Execution", JOptionPane.INFORMATION_MESSAGE);
                }
               //finally{
               //     driver.close();
               //}

        }

        public static void CloseBrowser() throws WebDriverException
        {
            driver.manage().deleteAllCookies();
            Set<String> allWindowHandles = driver.getWindowHandles();
            for(String handle : allWindowHandles){
                driver.switchTo().window(handle);
                driver.close();
            }
            
            //driver.quit();
        }

        public static boolean SendKeys_EventById(String ElementIdString, String ElementValue, String ElementReference) throws Exception
        {
            boolean respondresult = false;
            String errCatchMsg = "";
            String errMsg = "";
            
//            try{
                    if(!ElementReference.equals("null"))
                    {
                            //ElementIdString = ElementIdString + ElementReference;
                            Properties propObj = new Properties();
                            ExcelObj excelObj = new ExcelObj();

                            String value = "";
                            //value = excelObj.ReadToTemp(ElementReference); 
                            value = ewb.qa.tdd.GUI.ExcelObject.ReadToTemp(ElementReference);
                            ElementValue = value;
                    }

                    driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                    
                    boolean inputVisible = driver.findElement(By.id(ElementIdString)).isDisplayed();
                    boolean inputEnable = driver.findElement(By.id(ElementIdString)).isEnabled();

                    //Commented due to error 12/16/2020
                    if(inputVisible == true && inputEnable == true)
                    {
                            //String currentvalue = driver.findElement(By.id(ElementIdString)).getText().toString();
                            String currentvalue = driver.findElement(By.id(ElementIdString)).getAttribute("value");
                            String tagname = driver.findElement(By.id(ElementIdString)).getTagName();
                            
                            if(!currentvalue.equals("")){
                                driver.findElement(By.id(ElementIdString)).clear();
                            }
                            else if(!currentvalue.equalsIgnoreCase("")){
                                driver.findElement(By.id(ElementIdString)).clear();
                            }

                            if(ElementValue.equals("null")){
                                ElementValue = "";
                            }
                            else if(ElementValue.equalsIgnoreCase("null")){
                                ElementValue = "";
                            }
                            else if(ElementValue == "null"){
                                ElementValue = "";
                            }
                            //driver.findElement(By.id(ElementIdString)).clear();
                            driver.findElement(By.id(ElementIdString)).sendKeys(ElementValue);	
                            respondresult = true;
                    }
                    else{
                        respondresult = false;
                        errCatchMsg = "Element ID :" + ElementIdString + "\n" + "Error Message : " + "Element ID/ Xpath is not visible or enabled for Selenium." + "\n" +
                                "Action : " + "Check and verify element mapping or use other alternate element mapping.";
                        
                        errMsg = getEventErrorMsg();
                        if(errMsg == "" || errMsg.equals("")){
                            setEventErrorMsg(errCatchMsg);
                        }
                        else{
                            setEventErrorMsg(errCatchMsg);
                        }
                    }
                    return respondresult;
//            }
//            catch(Exception ex){
//                    //System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
//                    System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
//                    String LocalMsg = ex.getLocalizedMessage().toString();
//                    String StackTrace = ex.getStackTrace().toString();
//                    String Cause = ex.getCause().toString();
//                    String Message = ex.getMessage();
//                    errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
//
//                    errMsg = getEventErrorMsg();
//                    if(errMsg == "" || errMsg.equals("")){
//                        setEventErrorMsg(errCatchMsg);
//                    }
//                    else{
//                        setEventErrorMsg(errCatchMsg);
//                    }
//                    return false;                
//            }
        }

        public static boolean SendKeys_EventByXpath(String ElementXpathString, String ElementValue, String ElementReference) throws Exception
        {
            boolean respondresult = false;
            String errCatchMsg = "";
            String errMsg = "";
            
//            try{
                    if(!ElementReference.equals("null"))
                    {
                        //ElementIdString = ElementIdString + ElementReference;
                        Properties propObj = new Properties();
                        ExcelObj excelObj = new ExcelObj();

                        String value = "";
                        //value = excelObj.ReadToTemp(ElementReference); 
                        value = ewb.qa.tdd.GUI.ExcelObject.ReadToTemp(ElementReference);
                        ElementValue = value;
                    }
                    //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                    
                    driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                    //Commented due to error 12/16/2020
                    boolean inputVisible = driver.findElement(By.xpath(ElementXpathString)).isDisplayed();
                    boolean inputEnable = driver.findElement(By.xpath(ElementXpathString)).isEnabled();

                    //Commented due to error 12/16/2020
                    if(inputVisible == true && inputEnable == true)
                    {
                            //String currentvalue = driver.findElement(By.xpath(ElementXpathString)).getText().toString();
                            String currentvalue = driver.findElement(By.xpath(ElementXpathString)).getAttribute("value");
                            //String tagname = driver.findElement(By.xpath(ElementXpathString)).getTagName();

                            if(!currentvalue.equals("")){
                                driver.findElement(By.id(ElementXpathString)).clear();
                            }
                            else if(!currentvalue.equalsIgnoreCase("")){
                                driver.findElement(By.id(ElementXpathString)).clear();
                            }

                            if(ElementValue.equals("null")){
                                ElementValue = "";
                            }
                            else if(ElementValue.equalsIgnoreCase("null")){
                                ElementValue = "";
                            }
                            else if(ElementValue == "null"){
                                ElementValue = "";
                            }                    

                            //driver.findElement(By.xpath(ElementXpathString)).clear();
                            driver.findElement(By.xpath(ElementXpathString)).sendKeys(ElementValue);	
                            respondresult = true;
                    }
                    else{
                        respondresult = false;
                        errCatchMsg = "Element ID :" + ElementXpathString + "\n" + "Error Message : " + "Element ID/ Xpath is not visible or enabled for Selenium." + "\n" +
                                    "Action : " + "Check and verify element mapping or user other alternate element mapping.";
                        
                        errMsg = getEventErrorMsg();
                        if(errMsg == "" || errMsg.equals("")){
                            setEventErrorMsg(errCatchMsg);
                        }
                        else{
                            setEventErrorMsg(errCatchMsg);
                        }
                    }
                    return respondresult;
//            }
//            catch(Exception ex){
//                        //System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
//                        System.out.println("Element XPath :" + ElementXpathString + "\n" + ex.getMessage());
//                        String LocalMsg = ex.getLocalizedMessage().toString();
//                        String StackTrace = ex.getStackTrace().toString();
//                        String Cause = ex.getCause().toString();
//                        String Message = ex.getMessage();
//                        errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
//                        
//                        errMsg = getEventErrorMsg();
//                        if(errMsg == "" || errMsg.equals("")){
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        else{
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        return false;
//            }
        }


        public static boolean Click_EventById(String ElementIdString) throws Exception{
            boolean respondresult = false;
            String errCatchMsg = "";
            String errMsg = "";                
            
//                try
//                {
                        //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                        driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                        boolean searchElementSelected = driver.findElement(By.id(ElementIdString)).isSelected();
                        
                        //Commented due to error 12/16/2020
                        boolean searchElementPresence = driver.findElement(By.id(ElementIdString)).isDisplayed();
                        boolean searchElementEnabled = driver.findElement(By.id(ElementIdString)).isEnabled();

                        if(searchElementSelected == false)
                        {
                                //Commented due to error 12/16/2020
                                if(searchElementPresence == true && searchElementEnabled == true)
                                {
                                    driver.findElement(By.id(ElementIdString)).click();
                                    //Switchto_Default();
                                    respondresult = true;
                                }
                                else{
                                    respondresult = false;
                                    
                                    errCatchMsg = "Element ID :" + ElementIdString + "\n" + "Error Message : " + "Element ID/ Xpath is not visible or enabled for Selenium." + "\n" +
                                                                    "Action : " + "Check and verify element mapping or user other alternate element mapping.";

                                    errMsg = getEventErrorMsg();
                                    if(errMsg == "" || errMsg.equals("")){
                                        setEventErrorMsg(errCatchMsg);
                                    }
                                    else{
                                        setEventErrorMsg(errCatchMsg);
                                    }                                    
                                }
                        }
                        return respondresult;
//                }
//                catch(Exception ex)
//                {
//                        //System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
//                        System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
//                        String LocalMsg = ex.getLocalizedMessage().toString();
//                        String StackTrace = ex.getStackTrace().toString();
//                        String Cause = ex.getCause().toString();
//                        String Message = ex.getMessage();
//                        errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
//                        
//                        errMsg = getEventErrorMsg();
//                        if(errMsg == "" || errMsg.equals("")){
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        else{
//                            setEventErrorMsg(errCatchMsg);
//                        } 
//                        return false;
//                }

        }

        public static boolean Click_EventByXpath(String ElementXpathString) throws Exception{
                //boolean eventresult = true;
                boolean respondresult = false;
                String errCatchMsg = "";
                String errMsg = "";
                
//                try
//                {
                        //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                        driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                        //String elemType = driver.findElement(By.xpath(ElementXpathString)).getAttribute("type");
                        
                        boolean searchElementSelected = driver.findElement(By.xpath(ElementXpathString)).isSelected();

                                                    //Commented due to error 12/16/2020
                        boolean searchElementPresence = driver.findElement(By.xpath(ElementXpathString)).isDisplayed();
                        boolean searchElementEnabled = driver.findElement(By.xpath(ElementXpathString)).isEnabled();

                                //Commented due to error 12/16/2020
                                if(searchElementPresence == true && searchElementEnabled == true)
                                {
                                    if(searchElementSelected == false){
                                        driver.findElement(By.xpath(ElementXpathString)).click();
                                        //Switchto_Default();
                                        respondresult = true;                                        
                                    }
                                    else{
                                        respondresult = true;
                                    }
                                }
                                else{
                                    respondresult = false;
                                    errCatchMsg = "Element ID :" + ElementXpathString + "\n" + "Error Message : " + "Element ID/ Xpath is not visible or enabled for Selenium." + "\n" +
                                                "Action : " + "Check and verify element mapping or user other alternate element mapping.";

                                    errMsg = getEventErrorMsg();
                                    if(errMsg == "" || errMsg.equals("")){
                                        setEventErrorMsg(errCatchMsg);
                                    }
                                    else{
                                        setEventErrorMsg(errCatchMsg);
                                    }
                                }
                        //Switchto_Default();	
                        return respondresult;
//                }
//                catch(Exception ex)
//                {
//                        System.out.println("Element ID :" + ElementXpathString + "\n" + ex.getMessage());
//                        String LocalMsg = ex.getLocalizedMessage().toString();
//                        String StackTrace = ex.getStackTrace().toString();
//                        String Cause = ex.getCause().toString();
//                        String Message = ex.getMessage();
//                        errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
//                        
//                        errMsg = getEventErrorMsg();
//                        if(errMsg == "" || errMsg.equals("")){
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        else{
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        return false;
//                }
        }

        public static String Text_EventById(String ElementIdString){
                SetText_EventById(driver.findElement(By.id(ElementIdString)).getText().toString());
                //Switchto_Default();
                return GetText_EventById();
        }

        public static void SetText_EventById(String strBrowserMessage){
                strbrowsermessage = strBrowserMessage;
        }

        public static String GetText_EventById(){
                return strbrowsermessage;
        }

        public static String Text_EventByXpath(String ElementXpathString){
                SetText_EventByXpath(driver.findElement(By.xpath(ElementXpathString)).getText().toString());
                //Switchto_Default();
                return GetText_EventByXpath();
        }

        public static void SetText_EventByXpath(String strBrowserMessage){
                strbrowsermessage = strBrowserMessage;
        }

        public static String GetText_EventByXpath(){
                return strbrowsermessage;
        }

        public static boolean GetInputText_EventById(String ElementIdString, String ElementReference) throws Exception{
            boolean result = false;
            String errCatchMsg = "";
            String errMsg = "";
            
//            try{
                //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                
                //Commented due to error 12/16/2020
                boolean inputVisible = driver.findElement(By.xpath(ElementIdString)).isDisplayed();
                boolean inputEnable = driver.findElement(By.xpath(ElementIdString)).isEnabled();
                
                 if(inputVisible == true && inputEnable == true){
                     
                        String currentvalue = "";
                        currentvalue = driver.findElement(By.xpath(ElementIdString)).getAttribute("textContent");
                        int strLength = currentvalue.length();
                        if(strLength == 0){
                            currentvalue = driver.findElement(By.xpath(ElementIdString)).getAttribute("value");
                        }

                        ExcelObject.StoreToTemp(currentvalue, ElementReference,"");
                        result = true;
                 }
                return result;            
        }
        
        public static boolean GetInputText_EventByXpath(String ElementXpathString, String ElementReference) throws Exception{
            boolean result = false;
            String errCatchMsg = "";
            String errMsg = "";
            
//            try{
                //driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
                driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                
                //Commented due to error 12/16/2020
                boolean inputVisible = driver.findElement(By.xpath(ElementXpathString)).isDisplayed();
                boolean inputEnable = driver.findElement(By.xpath(ElementXpathString)).isEnabled();
                
                 if(inputVisible == true && inputEnable == true){
                     
                        String currentvalue = "";
                        currentvalue = driver.findElement(By.xpath(ElementXpathString)).getAttribute("textContent");
                        int strLength = currentvalue.length();
                        if(strLength == 0){
                            currentvalue = driver.findElement(By.xpath(ElementXpathString)).getAttribute("value");
                        }

                        ExcelObject.StoreToTemp(currentvalue, ElementReference,"");
                        result = true;
                 }
                return result;
//            }
//            catch(Exception ex){
//                        System.out.println("Element XPath :" + ElementXpathString + "\n" + ex.getMessage());
//                        String LocalMsg = ex.getLocalizedMessage().toString();
//                        String StackTrace = ex.getStackTrace().toString();
//                        String Cause = ex.getCause().toString();
//                        String Message = ex.getMessage();
//                        errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
//                        //
//                        
//                        errMsg = getEventErrorMsg();
//                        if(errMsg == "" || errMsg.equals("")){
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        else{
//                            setEventErrorMsg(errCatchMsg);
//                        }
//                        return result;                
//            }
        }
        
        public static boolean Select_EventById(String ElementIdString, int frameNum) throws Exception{
            try{
                /*
                List<WebElement> iframeElements = driver.findElements(By.tagName("frame"));
                int iframeCount = iframeElements.size();
                int iCtr = 0;
                int iframeValue = frameNum - 1;

                for(WebElement e: iframeElements)
                {
                        switch(iCtr)
                        {
                        case 0:
                                if(iframeValue == iCtr)
                                {
                                        driver.switchTo().frame(driver.findElement(By.xpath(ElementIdString)));
                                }
                                break;

                        case 1:
                                if(iframeValue == iCtr)
                                {
                                        JavascriptExecutor js = (JavascriptExecutor) driver;
                                        js.executeScript("document.getElementsByTagName('" + e.getTagName() + "')[" + iCtr + "].contentWindow.location.reload();");

                                        driver.switchTo().frame(driver.findElement(By.xpath(ElementIdString)));
                                }
                                break;
                        }
                        iCtr++;
                }
                boolSwitchto = true;
                */
                    List<WebElement> iframeElements = driver.findElements(By.tagName("frame"));
                    if(!iframeElements.isEmpty()){
                            int iframeCount = iframeElements.size();
                            int iCtr = 0;
                            int iframeValue = frameNum - 1;

                            for(WebElement e: iframeElements)
                            {
                                    switch(iCtr)
                                    {
                                    case 0:
                                            if(iframeValue == iCtr)
                                            {
                                                    driver.switchTo().frame(driver.findElement(By.xpath(ElementIdString)));
                                            }
                                            break;

                                    case 1:
                                            if(iframeValue == iCtr)
                                            {
                                                    JavascriptExecutor js = (JavascriptExecutor) driver;
                                                    js.executeScript("document.getElementsByTagName('" + e.getTagName() + "')[" + iCtr + "].contentWindow.location.reload();");

                                                    driver.switchTo().frame(driver.findElement(By.xpath(ElementIdString)));
                                            }
                                            break;
                                    }
                                    iCtr++;
                            }
                            boolSwitchto = true;   
                            return true;

                    }
                    else{
                        return false;
                    }
            }
            catch(Exception ex){
                    System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
                    String LocalMsg = ex.getLocalizedMessage().toString();
                    String StackTrace = ex.getStackTrace().toString();
                    String Cause = ex.getCause().toString();
                    String Message = ex.getMessage();
                    String errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;

                    String errMsg = getEventErrorMsg();
                    if(errMsg == "" || errMsg.equals("")){
                        setEventErrorMsg(errCatchMsg);
                    }
                    else{
                        setEventErrorMsg(errCatchMsg);
                    }                
                    return false;
            }

        }

        public static boolean Select_EventByXpath(String ElementXpathString, int frameNum) throws Exception{		
            try{
                List<WebElement> iframeElements = driver.findElements(By.tagName("frame"));
                if(!iframeElements.isEmpty()){
                        int iframeCount = iframeElements.size();
                        int iCtr = 0;
                        int iframeValue = frameNum - 1;

                        for(WebElement e: iframeElements)
                        {
                                switch(iCtr)
                                {
                                case 0:
                                        if(iframeValue == iCtr)
                                        {
                                                driver.switchTo().frame(driver.findElement(By.xpath(ElementXpathString)));
                                        }
                                        break;

                                case 1:
                                        if(iframeValue == iCtr)
                                        {
                                                JavascriptExecutor js = (JavascriptExecutor) driver;
                                                js.executeScript("document.getElementsByTagName('" + e.getTagName() + "')[" + iCtr + "].contentWindow.location.reload();");

                                                driver.switchTo().frame(driver.findElement(By.xpath(ElementXpathString)));
                                        }
                                        break;
                                }
                                iCtr++;
                        }
                        boolSwitchto = true;   
                        return true;

                }
                else{
                    return false;
                }

            }
            catch(Exception ex){
                        System.out.println("Element Xpath :" + ElementXpathString + "\n" + ex.getMessage());
                        String LocalMsg = ex.getLocalizedMessage().toString();
                        String StackTrace = ex.getStackTrace().toString();
                        String Cause = ex.getCause().toString();
                        String Message = ex.getMessage();
                        String errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
                        
                        String errMsg = getEventErrorMsg();
                        if(errMsg == "" || errMsg.equals("")){
                            setEventErrorMsg(errCatchMsg);
                        }
                        else{
                            setEventErrorMsg(errCatchMsg);
                        }
                        return false;                
            }

        }

        public static String Title_NewTab(){
                SetTitle_NewTab();
                //Switchto_Default();
                return GetTitle_NewTab();
        }

        public static void SetTitle_NewTab(){
                String strTitleNewTab = driver.getTitle().toString();
                strnewtabtitle = strTitleNewTab;
        }

        public static String GetTitle_NewTab(){
                return strnewtabtitle;
        }

        public static void NewTab_Switchto(){
                Set<String> tab = driver.getWindowHandles();
                int number_of_tabs = tab.size();
                int new_tab_index = number_of_tabs -1;
                driver.switchTo().window(tab.toArray()[new_tab_index].toString());

                //Commented due to error 12/16/2020
                //driver.manage().window().maximize();

        }
        
        public static boolean NewWindow_Select(String givenTitle){
            JFrame jFrame = new JFrame();
            try{
                String parent = driver.getWindowHandle();
                Set<String> allWindows = driver.getWindowHandles();
                int count = allWindows.size();

                for(String child:allWindows){
                    if(!parent.equalsIgnoreCase(child)){
                        String windowtitle = driver.switchTo().window(child).getTitle();
                        if(windowtitle.equals(givenTitle)){
                            driver.switchTo().window(child);
                        }
                    }
                }           
                return true;
            }
            catch(Exception ex){
                //System.out.println("Element Xpath :" + ElementXpathString + "\n" + ex.getMessage());
                String LocalMsg = ex.getLocalizedMessage().toString();
                String StackTrace = ex.getStackTrace().toString();
                String Cause = ex.getCause().toString();
                String Message = ex.getMessage();
                String errCatchMsg = "Element Type :" + "NewWindow_Select" + ", " + "Element Action :" + "Switch To" + "\n" + 
                        "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;

                String errMsg = getEventErrorMsg();
                if(errMsg == "" || errMsg.equals("")){
                    setEventErrorMsg(errCatchMsg);
                }
                else{
                    setEventErrorMsg(errCatchMsg);
                }
                return false;                 
            }

        }

        public static void NewTab_Close(){
                Set<String> tab = driver.getWindowHandles();
                int number_of_tabs = tab.size();
                int new_tab_index = number_of_tabs -1;
                driver.switchTo().window(tab.toArray()[new_tab_index].toString());
                driver.close();
                new_tab_index = new_tab_index - 1;
                driver.switchTo().window(tab.toArray()[new_tab_index].toString());
        }

        public static void Submit_EventById(String ElementIdString){
                driver.findElement(By.id(ElementIdString)).submit();
        }

        public static void Submit_EventByXpath(String ElementXpathString){
                driver.findElement(By.xpath(ElementXpathString)).submit();
        }

        public static void SubTab_SelectById(String ElementIdString){
                driver.findElement(By.id(ElementIdString)).click();
        }

        public static void SubTab_SelectByXpath(String ElementXpathString){
                driver.findElement(By.xpath(ElementXpathString)).click();
        }

        public static boolean DropdownBox_ById(String ElementIdString, String ElementValue, String ElementReference) throws Exception{
                boolean respondresult = false;
                String errCatchMsg = "";
                String errMsg = "";                
            
                try{
                        if(!ElementReference.equals("null"))
                        {
                                //ElementIdString = ElementIdString + ElementReference;
                                Properties propObj = new Properties();
                                ExcelObj excelObj = new ExcelObj();

                                String value = "";
                                //value = excelObj.ReadToTemp(ElementReference); 
                                value = ewb.qa.tdd.GUI.ExcelObject.ReadToTemp(ElementReference);
                                ElementIdString = ElementIdString + value;
                        }

                        driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                        String dropdownValue = driver.findElement(By.id(ElementIdString)).getText().toString();
                        if(dropdownValue != ElementValue)
                        {
                            boolean visible = driver.findElement(By.id(ElementIdString)).isDisplayed();
                            boolean enabled = driver.findElement(By.id(ElementIdString)).isEnabled();

                            if(visible == true && enabled == true){
                                Select drpElement = new Select(driver.findElement(By.id(ElementIdString)));
                                drpElement.selectByVisibleText(ElementValue);
                                respondresult = true;
                            }
                            else{
                                respondresult = false;
                                errCatchMsg = "Element ID :" + ElementIdString + "\n" + "Error Message : " + "Element ID/ Xpath is not visible or enabled for Selenium." + "\n" +
                                            "Action : " + "Check and verify element mapping or user other alternate element mapping.";

                                errMsg = getEventErrorMsg();
                                if(errMsg == "" || errMsg.equals("")){
                                    setEventErrorMsg(errCatchMsg);
                                }
                                else{
                                    setEventErrorMsg(errCatchMsg);
                                }
                            }
                        }
                        return respondresult;
                }
                catch(Exception ex){
                        System.out.println("Element ID :" + ElementIdString + "\n" + ex.getMessage());
                        String LocalMsg = ex.getLocalizedMessage().toString();
                        String StackTrace = ex.getStackTrace().toString();
                        String Cause = ex.getCause().toString();
                        String Message = ex.getMessage();
                        errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
                        
                        errMsg = getEventErrorMsg();
                        if(errMsg == "" || errMsg.equals("")){
                            setEventErrorMsg(errCatchMsg);
                        }
                        else{
                            setEventErrorMsg(errCatchMsg);
                        }
                        return false;                       
                }
                //System.out.println("Element ID :" + ElementIdString);
        }

        public static boolean DropdownBox_ByXpath(String ElementXpathString, String ElementValue, String ElementReference) throws Exception{	
                boolean respondresult = false;
                String errCatchMsg = "";
                String errMsg = "";                    
            
             try{
                        if(!ElementReference.equals("null"))
                        {
                            //ElementXpathString = ElementXpathString + ElementReference;
                            Properties propObj = new Properties();
                            ExcelObj excelObj = new ExcelObj();

                            driver.manage().timeouts().implicitlyWait(25, TimeUnit.SECONDS);
                            boolean visible = driver.findElement(By.xpath(ElementXpathString)).isDisplayed();
                            boolean enabled = driver.findElement(By.xpath(ElementXpathString)).isEnabled();

                            if(visible == true && enabled == true){
                                String value = "";
                                //value = excelObj.ReadToTemp(ElementReference); 
                                value = ewb.qa.tdd.GUI.ExcelObject.ReadToTemp(ElementReference);
                                ElementXpathString = ElementXpathString + value;         
                                respondresult = true;
                            }
                            else{
                                respondresult = false;
                                errCatchMsg = "Element ID :" + ElementXpathString + "\n" + "Error Message : " + "Element ID/ Xpath is not visible or enabled for Selenium." + "\n" +
                                            "Action : " + "Check and verify element mapping or user other alternate element mapping.";

                                errMsg = getEventErrorMsg();
                                if(errMsg == "" || errMsg.equals("")){
                                    setEventErrorMsg(errCatchMsg);
                                }
                                else{
                                    setEventErrorMsg(errCatchMsg);
                                }                                    
                            }
                        }

                        String dropdownValue = driver.findElement(By.xpath(ElementXpathString)).getText().toString();
                        if(dropdownValue != ElementValue)
                        {
                                Select drpElement = new Select(driver.findElement(By.xpath(ElementXpathString)));
                                drpElement.selectByVisibleText(ElementValue);			
                        }	
                        return true;
                }
                catch(Exception ex){
                        System.out.println("Element Xpath :" + ElementXpathString + "\n" + ex.getMessage());
                        String LocalMsg = ex.getLocalizedMessage().toString();
                        String StackTrace = ex.getStackTrace().toString();
                        String Cause = ex.getCause().toString();
                        String Message = ex.getMessage();
                        errCatchMsg = "Local Message :" + LocalMsg + "\n" + "Stack Trace :" + StackTrace + "\n" + "Cause :" + Cause + "\n" + "Error Message :" + Message;
                        
                        errMsg = getEventErrorMsg();
                        if(errMsg == "" || errMsg.equals("")){
                            setEventErrorMsg(errCatchMsg);
                        }
                        else{
                            setEventErrorMsg(errCatchMsg);
                        }
                        return false;                                
                }
                //System.out.println("Element XPath :" + ElementXpathString);
        }

        public static void Select_Window() throws Exception{
                String winHandleBefore = driver.getWindowHandle();
                for(String winHandle : driver.getWindowHandles()){
                    driver.switchTo().window(winHandle);
                }
                
        }        

        public static String TextMessage_ById(String ElementIdString){
                SetTextMessage_ById(ElementIdString);
                return GetTextMessage_ById();
        }

        public static void SetTextMessage_ById(String ElementIdString){
                strtextmessage = driver.findElement(By.id(ElementIdString)).getText().toString();
        }

        public static String GetTextMessage_ById(){
                return strtextmessage;
        }

        public static String TextMessage_ByXpath(String ElementXpathString){
                SetTextMessage_ByXpath(ElementXpathString);
                return GetTextMessage_ByXpath();
        }

        public static void SetTextMessage_ByXpath(String ElementXpathString){
                strtextmessage = driver.findElement(By.xpath(ElementXpathString)).getText().toString();
        }

        public static String GetTextMessage_ByXpath(){
                return strtextmessage;
        }

        public static void ValidateElement(String ElementType, String Action){

        }

        public static void Logout()
        {

        }

        public static void Switchto_Default(){
                        if(boolSwitchto == true)
                        {
                                boolSwitchto = false;
                                driver.switchTo().defaultContent();

                        }
        }
        
        public static void setEventErrorMsg(String givenValue){
            eventerrormsg = givenValue;
        }
        
        public static String getEventErrorMsg(){
            if(eventerrormsg == null){
                setEventErrorMsg("");
            }
            else if(eventerrormsg.equals("null")){
                setEventErrorMsg("");
            }
            return eventerrormsg;
        }
        
        public static void ForceStopSelenium(){
            JFrame jFrame = new JFrame();
            int response = -1;
            response = JOptionPane.showConfirmDialog(jFrame, "Are you sure you to stop Selenium?", "Execution Page - TDD", JOptionPane.YES_NO_OPTION);
            if(response == 1){
                driver.close();
                driver.quit();                
            }

        }
        
        private static WebDriver globalWebDriver;
        
        public void setWebDriver(WebDriver givenWebDriver){
            globalWebDriver = givenWebDriver;
        }
        
        public WebDriver getWebDriver(){
            return globalWebDriver;
        }
        
        public static String ListWebElement_ById(String givenElementId, String givenTagName, String givenLabel){
            //boolean respondresult = false;
            String newElementId = "";
            try{
                List<WebElement> lblelement = driver.findElements(By.tagName("label"));
                for(WebElement elem: lblelement){
                    if(elem.getText().contains(givenLabel)){
                        newElementId = lblelement.toString();
                    }
                }
                //getText = "Arrangement"
                //Arrangement for == givenElementId 
                //givenElementId = Arrangement for
                //respondresult = true;
                return newElementId;
            }
            catch(Exception ex){
                
                return newElementId;
            }
        }
        
        public static boolean ListWebElement_ByXpath(String givenElementXpath, String givenTagName, String givenLabel, String givenType){
            boolean respondresult = false;
            //String newElementXpath = "";
            try{
                //List<WebElement>
                //getText = "Arrangement"
                //Arrangement for == givenElementId 
                //givenElementId = Arrangement for
                List<WebElement> lblelement = driver.findElements(By.tagName(givenTagName));
                for(WebElement elem: lblelement){
                    String elemId = elem.getAttribute("id");
                    String elemName = elem.getAttribute("name");
                    String elemType = elem.getAttribute("type");
                    String elemValue = elem.getAttribute("value");                    
                    
                    System.out.println("Element Id: " + elemId + "\n" + 
                            "Element Name: " + elemName + "\n" +
                            "Element Type: " + elemType + "\n" +
                            "Element Value: " + elemValue);
                    
                    if(elemType.equals(givenType)){
                        if(elemValue.equals(givenLabel)){
                            //System.out.println("Found element");
                            boolean elemVisible = elem.isDisplayed();
                            boolean elemEnable = elem.isEnabled();
                            boolean elemSelected = elem.isSelected();
                            if(elemVisible && elemEnable && !elemSelected){
                                elem.click();
                                respondresult = true;
                            }
                        }
                    }
                }                
                return respondresult;
            }
            catch(Exception ex){
                return respondresult;
            }
        }
        
        public Response execute(Command command) throws IOException{
            Response response = null;
            
            return response;
        }
        
        public static void getCurrentWindow(String givenTitle) throws Exception{
            Set<String> allWindows = driver.getWindowHandles();
            int count = allWindows.size();

            for(String child:allWindows){
                String windowtitle = driver.switchTo().window(child).getTitle();
                if(windowtitle.equals(givenTitle)){
                    driver.switchTo().window(child);
                }
            }    
        }
}
