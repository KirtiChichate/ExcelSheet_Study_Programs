package excelIntro;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelPrintWholeSheet {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");

	    Sheet Getsheet = WorkbookFactory.create(ObjFile).getSheet("Intro2");
	    
	  int lastrow = Getsheet.getLastRowNum();
	  
	 int lastcell = Getsheet.getRow(0).getLastCellNum()-1; // for getlastcellnum-1 -->java.lang.NullPointerException---->this exception shows becz when we are not taken lastcell-1 then it count from 1th index not from zero index 
	 
	 for(int i=0; i<=lastrow; i++)
	 {
		 for(int j=0; j<=lastcell; j++)
		 {
			String allrowcell = Getsheet.getRow(i).getCell(j).getStringCellValue();
			System.out.print(allrowcell+" ");
		 }
		 System.out.println();
	 }
	 
	 
	 
	 
	}

}
