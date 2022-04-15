package excelIntro;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUsingLastRow_LastCell {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");  

		Sheet Getsheet = WorkbookFactory.create(ObjFile).getSheet("IntroExcel");
		
		//this getlastrownum will print till last row
		int LastRow = Getsheet.getLastRowNum();
		System.out.println(LastRow);
		for(int i=0; i<=LastRow; i++)
		{
			String printtilllastrow = Getsheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println(printtilllastrow);
		}
		System.out.println("############################################");
		
		
		
		//this getlastcolumnnum will print till last column
		short LastCell = Getsheet.getRow(0).getLastCellNum();
		System.out.println(LastCell);
    	int totalcell = LastCell-1;    // java.lang.NullPointerException---->this exception shows becz when we are not taken lastcell-1 then it count from 1th index not from zero index 
		for(int i=0; i<=totalcell; i++)
		{
			String printtilllastcell = Getsheet.getRow(0).getCell(i).getStringCellValue();
			System.out.print(printtilllastcell+" ");
		}
// exception---->  java.lang.IllegalStateException--->  Cannot get a STRING value from a NUMERIC cell becz in last in excel it have 
	}

}
