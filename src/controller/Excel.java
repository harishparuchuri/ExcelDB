package controller;
import java.io.File;  
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Iterator;  
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import utility.ConnectionManager;  
public class Excel  
{  

	public static void main(String[] args)   
	{  
		try  
		{  
			//Excel sheet inport
			File file = new File("Employee.xlsx");   
			//Another file
			//File file = new File("KLU.xlsx"); 

			//Geting table name from sheet name
			String TableName = file.getName().split("\\.")[0];




			FileInputStream fis = new FileInputStream(file);    


			@SuppressWarnings("resource")

			//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis); 

			//creating a Sheet object to retrieve object  
			XSSFSheet sheet = wb.getSheetAt(0);    

			XSSFRow rowz = sheet.getRow(0);
			XSSFRow rowo = sheet.getRow(1);

			// Create a List to store the header data and datatype
			ArrayList<String> headerData = new ArrayList<>();
			ArrayList<String> DataType = new ArrayList<>();


			//Geting heading names 
			for (Cell cell : rowz) {
				switch (cell.getCellTypeEnum()) {
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						DataFormatter dataFormatter = new DataFormatter();
						java.util.Date now=cell.getDateCellValue();
						headerData.add(dataFormatter.formatCellValue(cell));

					} else {
						headerData.add(String.valueOf(cell.getNumericCellValue()));
					}
					break;
				case STRING:
					headerData.add(cell.getStringCellValue());

					break;
				case BOOLEAN:
					headerData.add(String.valueOf(cell.getBooleanCellValue()));

					break;
				default:

					break;
				}

			}



			//GEting heading names completed//

			//geting heading datatype//

			for (Cell cell : rowo) {
				switch (cell.getCellTypeEnum()) {
				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						DataFormatter dataFormatter = new DataFormatter();

						DataType.add("timestamp NOT NULL");
					} else {

						DataType.add("Number");

					}
					break;
				case STRING:

					DataType.add("Varchar(255)");
					break;
				case BOOLEAN:

					DataType.add("NUMBER(1)");
					break;
				default:

					break;
				}
			}


			//Geting heading datatype completed//





			//Generating query to create table and  insert data into table
			String query="";
			String Insert_Data= "";
			String qmark=" VALUES ( ";

			Iterator i = headerData.iterator();
			Iterator j = DataType.iterator();

			while (i.hasNext() && j.hasNext()) {

				String coloum=(String) i.next();
				Insert_Data+=coloum+" ,";
				query+=coloum+" "+j.next()+",";

				qmark+=" ?,";


			}
			query=query.substring(0,query.length()-1);
			Insert_Data=Insert_Data.substring(0,Insert_Data.length()-1).toUpperCase();
			qmark=qmark.substring(0,qmark.length()-1);




			String Create_table="CREATE TABLE "+TableName+"("+query+")";
			String Insert_Table="INSERT INTO "+TableName+" ("+Insert_Data+")"+qmark+")";




			Statement stmt = null;
			Connection con=ConnectionManager.getConnection();
			try {
				stmt = con.createStatement();
				stmt.executeUpdate(Create_table);
			} catch (SQLException e) {
				System.out.println(e);
			} finally {
				if (stmt != null) { stmt.close(); }
			}
			System.out.println(TableName+" table created\n");
			System.out.println("Insert Query Generated\n");


			PreparedStatement statement = con.prepareStatement(Insert_Table);  

			int a=1;



			//iterating over excel file 


			Iterator<Row> itr = sheet.iterator();   

			itr.next();

			while (itr.hasNext())                 
			{  

				Row row = itr.next(); 

				//iterating over each column  


				Iterator<Cell> cellIterator = row.cellIterator();   
				a=0;

				while (cellIterator.hasNext())   
				{  
					for(String type:DataType) {
						Cell cell = cellIterator.next();  




						++a;

						if(type=="Number")
						{
							int num=(int) cell.getNumericCellValue();
							statement.setInt(a,num);

						}
						else if(type=="Varchar(255)")
						{

							statement.setString(a,cell.getStringCellValue() );


						}
						else if(type=="timestamp NOT NULL")
						{

							java.util.Date enrollDate =  cell.getDateCellValue();
							statement.setTimestamp(a, new Timestamp(enrollDate.getTime()));
						}

						else
						{
							break;
						}

					} 


				}
				statement.addBatch();


			} 
			statement.executeBatch();
			wb.close();

			System.out.println("Query executed\n");
			System.out.println("data stored in table sucesfully");
			con.commit();
			con.close();

		} 
		catch(Exception e)  
		{  
			e.printStackTrace();  
		}  
	}  
}  