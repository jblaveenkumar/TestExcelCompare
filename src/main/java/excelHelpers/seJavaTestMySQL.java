// Switch from a Top Window To Frame
package excelHelpers;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;

public class seJavaTestMySQL {
	static LinkedHashMap lhMap;
	static ArrayList list;
	
	public static void main(String[] args) throws ClassNotFoundException, SQLException {
		// Connect with mysql jdbc driver
		Class.forName("com.mysql.jdbc.Driver"); 
		System.out.println("Driver Loaded");
		// Connect with mysql database
		Connection con=DriverManager.getConnection("jdbc:mysql://127.0.0.1:3306/sakila","root","pass");
		System.out.println("Connection Establised");
		// Create object of "Statement" using createStatement() method 
		Statement smt = con.createStatement();
		// Execute the the query & results will be saved in ResultSet interface   
		ResultSet rs = smt.executeQuery("SELECT * FROM sakila.actor");
		//loop for fetching all records in database table table
		ResultSetMetaData rsmd=rs.getMetaData();  
		int colSize=rsmd.getColumnCount();
		System.out.println(colSize);
		
		rs.last();
		int rowSize=rs.getRow();
		System.out.println(rowSize);
		ExcelUtility.createXLWorkbook("./ComparisonOutput\\workbook1.xlsx","sheet1" , rowSize+1);
		ExcelUtility.openXLWorkbook("./ComparisonOutput\\workbook1.xlsx");
		
		for(int cl=1;cl<=colSize;cl++) {
			System.out.println(rsmd.getColumnLabel(cl));
			ExcelUtility.setXLCellValue("sheet1", 0, cl-1, rsmd.getColumnLabel(cl));
		}
		list = new ArrayList(rowSize);
		int r=1;
		rs.beforeFirst();
		while(rs.next()) {
			for(int cl=1;cl<=colSize;cl++) {
				lhMap = new LinkedHashMap(colSize);
				System.out.println("r="+r);
				System.out.println(rs.getString(cl));
				lhMap.put(rsmd.getColumnName(cl),rs.getObject(cl));
				ExcelUtility.setXLCellValue("sheet1",r , cl-1, rs.getString(cl));
				list.add(lhMap);
			}
			System.out.println(lhMap);
			r++;
		}
		ExcelUtility.writeAndCloseXL("./ComparisonOutput\\workbook1.xlsx");
		System.out.println(lhMap);
		System.out.println(list);
//		while(rs.next())
//		{
//			String actor_id = rs.getString("actor_id");
//			System.out.println("actor_id - "+actor_id);
////			ExcelUtility.setXLCellValue("sheet1", 1, 0,actor_id );
//			String firstName = rs.getString("first_Name");
//			System.out.println("first_Name "+firstName);
//		}
	}
} 
