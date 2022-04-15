package excelUsing_AllDataType;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelAllDataTypeUsingFORloop {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");  
		
		Sheet Getsheet = WorkbookFactory.create(ObjFile).getSheet("AllDataType");
		
		int LastRow = Getsheet.getLastRowNum();
		
		int LastCell = Getsheet.getRow(0).getLastCellNum()-1;
		
		for(int i=0; i<=LastRow; i++)
		{
			for(int j=0; j<=LastCell; j++)	
			{
				
			Cell infoRowCell = Getsheet.getRow(i).getCell(j);
			CellType Type = infoRowCell.getCellType();
			
			if(Type==CellType.STRING)
			{
				String outputString = infoRowCell.getStringCellValue();
				System.out.print(outputString+" ");
			}
			
			else if(Type==CellType.NUMERIC)
			{
				double outputNumeric = infoRowCell.getNumericCellValue();
				System.out.print(outputNumeric+" ");
			}
			
			else if(Type==CellType.BOOLEAN)
			{
			    boolean outputBoolean = infoRowCell.getBooleanCellValue();
			    System.out.print(outputBoolean+" ");
			}
				
			}
			
			System.out.println();
		}
		
		
		
		
		

	}

}
