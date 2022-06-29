package com.example.ExcelAutomator;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class SQLUtils {

    private static final Logger LOGGER = Logger.getLogger(SQLUtils.class.getName());

    public static void closeQuietly(Connection connection){
        try {
            if (connection != null){
                connection.close();
                System.out.println("connection closed properly");
            }
        } catch (SQLException e) {
            LOGGER.log(Level.SEVERE, e.toString(), e);
            System.out.println("exception on connection close");
        }
    }

    public static void closeQuietly(PreparedStatement statement){
        try {
            if (statement != null){
                statement.close();
                System.out.println("statement closed properly");
            }
        } catch (SQLException e){
            LOGGER.log(Level.SEVERE, e.toString(), e);
            System.out.println("exception on statement close");
        }
    }

    public static void closeQuietly(ResultSet resultSet){
        try {
            if (resultSet != null){
                resultSet.close();
                System.out.println("Resultset properly closed");
            }
        } catch (SQLException e){
            LOGGER.log(Level.SEVERE, e.toString(), e);
            System.out.println("Exception on resultset close");
        }
    }

}
