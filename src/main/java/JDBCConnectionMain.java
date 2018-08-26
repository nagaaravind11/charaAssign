import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
/**
 *
 * @author sqlitetutorial.net
 */
public class JDBCConnectionMain {
 
    /**
     * Connect to the test.db database
     * @return the Connection object
     */
    private Connection connect() {
        // SQLite connection string
    	  String url = "jdbc:sqlite:/Users/nabalasubramania/Desktop/SQL/sqlite-tools-osx-x86-3240000/test";
        Connection conn = null;
        try {
            conn = DriverManager.getConnection(url);
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
        return conn;
    }
 
    
    /**
     * select all rows in the warehouses table
     * @throws IOException 
     */
    public void selectAll() throws IOException{
    	for(String table: returnTables() )
    	{
        String sql = "SELECT * FROM  " + table;
        
        try (Connection conn = this.connect();
             Statement stmt  = conn.createStatement();
             ResultSet rs    = stmt.executeQuery(sql)){
        	ResultSetMetaData rsmd = rs.getMetaData();
            int cols = rsmd.getColumnCount();
           List<String> colNames =new ArrayList<>();
            System.out.println("\n");
            for (int i=1;i<=cols;i++) {
            	colNames.add( rsmd.getColumnName(i));
            	 System.out.print(rsmd.getColumnName(i) +"  " );
              }
            List<List<String>> colValuesList =new ArrayList<>();
            while (rs.next()) {
            	  List<String> colValues =new ArrayList<>();
               System.out.println("\n");
                for (String colName :colNames) {
                    String colValue= rs.getString(colName);
                    colValues.add(colValue);
                    System.out.print(colValue +" ");
                    }
                colValuesList.add(colValues);
            }
            ExcelWriter.addToExcel(colNames ,colValuesList,"test","chera.xlsx");
        } catch (SQLException e) {
            System.out.println(e.getMessage());
        }
    }
    }
   
 public List<String> returnTables(){
 	
    List<String> tables =new ArrayList<>();
    tables.add("COMPANY");
    tables.add("DEPARTMENT");
    return tables;
    
 }
   
    /**
     * @param args the command line arguments
     * @throws InterruptedException 
     * @throws IOException 
     */
    public static void main(String[] args) throws InterruptedException, IOException {
        JDBCConnectionMain app = new JDBCConnectionMain();
     app.selectAll();
    }
 
}