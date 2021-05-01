/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ewb.qa.tdd;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import javax.net.ssl.SSLContext;

import com.microsoft.sqlserver.jdbc.SQLServerDataSource;

public class SQLObj {

	public static Connection ConnToDB() throws Exception
	{
                        Connection conn = null;
                        boolean connFlag = false;
                        String DBName = "";
                        //Dev Environment
                        
//                        String connectionUrl = "jdbc:sqlserver://172.28.20.101:40889;databaseName=TDD;loginTimeout=30;";
//                        String username = "rbgmisuser";
//                        String userpassword = "p@ssw0rd";
//                        DBName = "Dev Environment";
                        
                        
                        //NEWSIT Environment
                        
                        //String connectionUrl = "jdbc:sqlserver://T24SW12IDVM03,40889;databaseName=TDD;loginTimeout=30;";
                        String connectionUrl = "jdbc:sqlserver://172.25.93.55:40889;databaseName=TDD;loginTimeout=30;";
                        String username = "qatdduser";
                        String userpassword = "p@ssw0rd";
                        DBName = "SIT Environment";                        
                        
                        try
                        {
                            Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
                            conn = DriverManager.getConnection(connectionUrl, username, userpassword);
                            connFlag = true;
                        }
                        catch(Exception ex)
                        {
                            connFlag = false;
                            System.out.println("Stack Trace: " + ex.getStackTrace() + "\n" +
                                    "Message: " + ex.getMessage() + "\n" + "Cause: " + ex.getCause() + "\n" + "Localize Message: " + ex.getLocalizedMessage());
                        }

                        if(connFlag)
                        {
                            System.out.println("Connected Successful");
                        }
                        else
                        {
                            System.out.println("Connection Failed");
                        }
                        
                        setDBConnName(DBName);
                        return conn;
	}
	
	private static String globalDBConnName;
        
                  public static void setDBConnName(String givenConnName){
                      globalDBConnName = givenConnName;
                  }
                  
                  public static String getDBConnName(){
                      return globalDBConnName;
                  }
                  
	public static ResultSet SQLExecuteCommand(String command, String commandtype) throws Exception
	{
                        SQLServerDataSource ds = new SQLServerDataSource();
                        ResultSet rs = null;

                        //int iArrCount = arrParam.size();
                        String field = "";
                        String value = "";

                        try (Connection conn = ConnToDB();)
                        {
                                switch(commandtype)
                                {
                                case "PrepareStatement":
                                        PreparedStatement pStmt = null;
                                        pStmt = conn.prepareStatement(command);
                                        rs = pStmt.executeQuery();
                                        break;

                                case "PrepareCall":
                                        CallableStatement pCall = null;
                                        pCall = conn.prepareCall(command);
                                        pCall.execute();
                                        rs = pCall.getResultSet();
                                        break;
                                }

                        }
                        catch(Exception ex)
                        {
                                System.out.println(ex.getLocalizedMessage());
                        }

                        return rs;
			
	}
}
