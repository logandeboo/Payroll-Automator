package com.example.ExcelAutomator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.sql.*;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;

public class MKSConnect {

    Connection con;
    PreparedStatement statement;
    ResultSet qResult;
    ArrayList<XSSFWorkbook> workbooks;

    private static final Logger LOGGER = Logger.getLogger(SQLUtils.class.getName());


    public ArrayList<XSSFWorkbook> handleQuery(LocalDate start, LocalDate end, int goldRate, int platRate, String path) {

        String connectionUrl = //REDACTED//
        String query =  "exec SubContractBilling";


        try{
            con = DriverManager.getConnection(connectionUrl);
            statement = con.prepareStatement(query,ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
            System.out.println("Query start");
            qResult = statement.executeQuery();
            ParseData parser = new ParseData();
            System.out.println("Query end, begin parse");
            workbooks = parser.parseMaster(qResult, start, end, goldRate, platRate);
            System.out.println("End parse");

        } catch (SQLException | IOException e) {
            LOGGER.log(Level.SEVERE, e.toString(), e);
            System.out.println(e);
            //launch error window and tell user to retry
        } finally {
            SQLUtils.closeQuietly(qResult);
            SQLUtils.closeQuietly(statement);
            SQLUtils.closeQuietly(con);
        }

        return workbooks;
    }






}
