package utility;

import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.util.Properties;

public class ConnectionManager {
	
	public static Connection getConnection() throws Exception {
		

		Properties p=loadPropertiesFile();
		{
				
				String driver = p.getProperty("driver");
				String url= p.getProperty("url");
				String user=p.getProperty("username");
				String password=p.getProperty("password");
				 Class.forName(driver);
				 Connection con;
				 con=DriverManager.getConnection(url,user,password);
				 
				 return con;
				 

			}
		
	}
	
	public static Properties loadPropertiesFile() throws Exception {
		Properties prop = new Properties();	
		InputStream in = ConnectionManager.class.getClassLoader().getResourceAsStream("jdbc.properties");
		prop.load(in);
		in.close(); 
		return prop;
	}
}

