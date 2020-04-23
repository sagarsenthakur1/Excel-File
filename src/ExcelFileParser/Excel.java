package ExcelFileParser;

import java.io.IOException;
import java.sql.*;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Excel {
	 public static void main(String[] args) {
	        String jdbcURL = "jdbc:mysql://localhost:3306/employee?autoReconnect=true&useSSL=false";
	        String username = "root";
	        String password = "Ocean@123";
	 
	        String excelFilePath = "src\\Book1.xlsx";
	 
	        int batchSize = 20;
	 
	        Connection connection = null;
	 
	        try {
	            long start = System.currentTimeMillis();
	             
	            FileInputStream inputStream = new FileInputStream(excelFilePath);
	 
	            Workbook workbook = new XSSFWorkbook(inputStream);
	 
	            Sheet firstSheet = workbook.getSheetAt(0);
	            Iterator<Row> rowIterator = firstSheet.iterator();
	 
	            connection = DriverManager.getConnection(jdbcURL, username, password);
	            connection.setAutoCommit(false);
	            DatabaseMetaData dbm = connection.getMetaData();
	         // check if "employee" table is there
	         ResultSet tables = dbm.getTables(null, null, "emp", null);
	         if (!tables.next()) {
	        	 Statement stmt=connection.createStatement();
		          stmt.executeUpdate("create table emp (id int not null,name varchar(25))");
	         }
	           
	            String sql = "INSERT INTO emp (id, name) VALUES (?, ?)";
	            PreparedStatement statement = connection.prepareStatement(sql);    
	             
	            int count = 0;
	             
	            rowIterator.next(); // skip the header row
	             
	            while (rowIterator.hasNext()) {
	                Row nextRow = rowIterator.next();
	                Iterator<Cell> cellIterator = nextRow.cellIterator();
	 
	                while (cellIterator.hasNext()) {
	                    Cell nextCell = cellIterator.next();
	 
	                    int columnIndex = nextCell.getColumnIndex();
	 
	                    switch (columnIndex) {
	                    case 0:
	                        int id = (int) nextCell.getNumericCellValue();
	                        statement.setInt(1, id);
	                        break;
	                    case 1:
	                        String name=nextCell.getStringCellValue();
	                        statement.setString(2, name);
	                        break;
	             
	                    }
	 
	                }
	                 
	                statement.addBatch();
	                 
	                if (count % batchSize == 0) {
	                    statement.executeBatch();
	                }              
	 
	            }
	 
	            workbook.close();
	             
	            // execute the remaining queries
	            statement.executeBatch();
	  
	            connection.commit();
	            connection.close();
	             
	            long end = System.currentTimeMillis();
	            System.out.printf("Import done in %d ms\n", (end - start));
	             
	        } catch (IOException ex1) {
	            System.out.println("Error reading file");
	            ex1.printStackTrace();
	        } catch (SQLException ex2) {
	            System.out.println("Database error");
	            ex2.printStackTrace();
	        }
	 
	    }

}
