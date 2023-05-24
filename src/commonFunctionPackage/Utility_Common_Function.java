package commonFunctionPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utility_Common_Function 
{
	public static void evidenceFileCreator(String fileName, String requestBody, String responseBody) throws IOException 
	{
		File newfile=new File("C:\\Users\\hp\\Desktop\\testing\\rest assured\\" +fileName+".txt");
		System.out.println("a new text file created to record request and response of the API :" +newfile.getName());
		
		FileWriter datawrite=new FileWriter(newfile);
		datawrite.write("request Body :" +requestBody+"\n\n");
		datawrite.write("response Body :" +responseBody);
		datawrite.close();
		System.out.println("request body and response body are saved in : " +newfile.getName());
		
	}
	
	public static ArrayList<String> readDataExcel(String sheetName, String testCaseName ) throws IOException
	{
		ArrayList<String> arrayData=new ArrayList<String>();
		
		//step 1: create the object of file input stream		
		FileInputStream fis=new FileInputStream("C:\\Users\\hp\\Desktop\\testing\\rest assured\\Book1.xlsx");
		
		//step 2: access the excel file
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		
		//step 3: access the sheet name
		int countOfSheet=workbook.getNumberOfSheets();
				for(int i=0;i<countOfSheet;i++)
				{
					String fileSheetName=workbook.getSheetName(i);
							if(fileSheetName.equalsIgnoreCase(sheetName))
							{
								//step 4: access the row from where data is suppose to fetch								
								XSSFSheet sheet=workbook.getSheetAt(i);
								
								Iterator<Row> rows=sheet.iterator();
								//Row r=rows.next();
								while(rows.hasNext())
								{
									Row r1=rows.next();
							
									if(r1.getCell(0).getStringCellValue().equalsIgnoreCase(testCaseName))
									{
										Iterator<Cell> cellvalues=r1.cellIterator();
												while(cellvalues.hasNext())
												{
													String testData=cellvalues.next().getStringCellValue();
															arrayData.add(testData);
												}
									}
								}
							}
				}
				workbook.close();
			return arrayData;
	}
}
	