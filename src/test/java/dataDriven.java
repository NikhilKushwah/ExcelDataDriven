import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	
	public ArrayList<String> getData(String testcaseName) throws IOException
	{
		
		//Now storing it into arraylist
		ArrayList<String> a= new ArrayList<String>();
		
	FileInputStream fis=new FileInputStream("C:\\Users\\nikhil.kushwah\\Desktop\\Demodata.xlsx");
	XSSFWorkbook workbook=new XSSFWorkbook(fis);
	
	
	int sheets=workbook.getNumberOfSheets();
	for(int i=0; i<sheets; i++)
	{
		if(workbook.getSheetName(i).equalsIgnoreCase("Data"))
		{
			XSSFSheet sheet=workbook.getSheetAt(i);
			
			//Identify Testcase column by scanning the entire 1st row
			//With the help of iterator we can move to next cell
			
			Iterator<Row> rows= sheet.iterator(); //this will check in the row in the sheet 
			Row Firstrow= rows.next();
			Iterator<Cell> ce=  Firstrow.cellIterator(); //This will check cell in the row
			int k=0; // initializing this value for loop
			int column=0;
			
			while(ce.hasNext())
			{
				Cell value=ce.next();
				if(value.getStringCellValue().equalsIgnoreCase("Testcase"))
				{
					column=k;
				}
				k++;
			}
			System.out.println(column);
			
			 //now will be fetching purchase testcase row in Testcase column
			while(rows.hasNext())
			{
				Row r=rows.next();
				if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
				{
					//after you grab purchase testcase row n pull all the data of that row 
					// with the help of iterator will be getting the data 
					Iterator<Cell> cv=r.cellIterator();
					while(cv.hasNext())
					{
						// this data storing into arraylist
						Cell c=cv.next();
						if(c.getCellType()==CellType.STRING)
						{
							//this will add the string value
						a.add(c.getStringCellValue());
						}
						else
						{
							//this will add Numberic value
							a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
						}
					}
						
				}
				
			}
		} 
	}
	return a;
	
	
}

	
	public static void main(String[] args) throws IOException 
	{
		
	}

}
 