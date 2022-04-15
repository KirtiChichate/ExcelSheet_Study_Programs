package excelIntro;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProgram2 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");   
		
//		creat file and also get the sheet from excel and store it in one ref variable
		
	    Sheet FileCreate = WorkbookFactory.create(ObjFile).getSheet("IntroExcel");
	
	//call all rows and all cell one by one
        String String00 = FileCreate.getRow(0).getCell(0).getStringCellValue();
        String String10 = FileCreate.getRow(1).getCell(0).getStringCellValue();
        String String20 = FileCreate.getRow(2).getCell(0).getStringCellValue();
        String String30 = FileCreate.getRow(3).getCell(0).getStringCellValue();
        
//        String String60 = FileCreate.getRow(6).getCell(0).getStringCellValue(); excel cant read null value it shows exception 
        // exception----> java.lang.NullPointerException
        
        System.out.println(String00);
        System.out.println(String10);
        System.out.println(String20);
        System.out.println(String30);
        
 //       System.out.println(String60);   exception      exception       exception       exception
     
	}

}
