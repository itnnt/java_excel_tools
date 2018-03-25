package main.utils;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class SqliteJdbcReaderWriter {
	 /** 
     * Connect to a sqlite database 
	 * @dbfile sqlite database file path
	 * @return Connection
     */  
    public Connection connect(String dbfile) {  
        Connection conn = null;  
        try {  
            // db parameters  
            String url = String.format("jdbc:sqlite:%s", dbfile);  
            // create a connection to the database  
            conn = DriverManager.getConnection(url);  
        } catch (SQLException e) {  
            System.out.println(e.getMessage());  
        } finally {  
            try {  
                if (conn != null) {  
                    conn.close();  
                }  
            } catch (SQLException ex) {  
                System.out.println(ex.getMessage());  
            }  
        }
        return conn;
    }  
    
    /**
     * Execute select sql command
     * @param dbfile
     * @param sql
     * @return
     */
    public ResultSet select(String dbfile, String sql) {  
    	ResultSet  rs = null;
        try{  
            Connection conn = this.connect(dbfile);  
            PreparedStatement pstmt = conn.prepareStatement(sql);  
            rs = pstmt.executeQuery(sql);  
        } catch (SQLException e) {  
            System.out.println(e.getMessage());  
        } 
        return rs;
    } 
    /** 
     * @param args the command line arguments 
     */  
    public static void main(String[] args) {  
    	SqliteJdbcReaderWriter sqliteJdbcRW = new SqliteJdbcReaderWriter();
    	sqliteJdbcRW.connect("D:\\workspace_data_processing\\gvl_data_utilities\\KPI_PRODUCTION\\dbbk\\main_database.db");  
    }
}
