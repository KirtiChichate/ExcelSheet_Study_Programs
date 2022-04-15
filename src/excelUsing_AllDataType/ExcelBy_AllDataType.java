package excelUsing_AllDataType;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelBy_AllDataType {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {


		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");   

		Sheet GetFile = WorkbookFactory.create(ObjFile).getSheet("AllDataType");
		//with the help of getcelltype we can perform all data type like STRING, NUMERIC, BOOLEAN
		
		CellType output10 = GetFile.getRow(1).getCell(0).getCellType();
		CellType output11 = GetFile.getRow(1).getCell(1).getCellType();
		CellType output12 = GetFile.getRow(1).getCell(2).getCellType();
		CellType output15 = GetFile.getRow(1).getCell(5).getCellType();
		CellType output16 = GetFile.getRow(1).getCell(6).getCellType();
		CellType output17 = GetFile.getRow(1).getCell(7).getCellType();
		
		System.out.println(output10);
		System.out.println(output11);
		System.out.println(output12);
		System.out.println(output15);
		System.out.println(output16);
		System.out.println(output17);
		
		
		boolean value1 = GetFile.getRow(1).getCell(7).getBooleanCellValue();
		 double value2 = GetFile.getRow(1).getCell(6).getNumericCellValue();
		  String value3 = GetFile.getRow(1).getCell(2).getStringCellValue();
		  
		  System.out.println(value1);
		  System.out.println(value2);
		  System.out.println(value3);
	}

}
