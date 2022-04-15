package excelIntro;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUsingForLoop {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");  

		Sheet Getsheet = WorkbookFactory.create(ObjFile).getSheet("IntroExcel");
		
		//suppose we want for data from zeroth column of all rows then with the help of FORLOOP we can print
		//but it can print till we finalize the i value so it is static type of nature
		for(int i=0; i<=4; i++)   //here loop start from 0 as we take 0 index first in excel
		{
			String Strongio = Getsheet.getRow(i).getCell(0).getStringCellValue();
			System.out.println(Strongio);
		}
		
		
		//same we can do for coulmn also we take row constant and vary the column
		for(int i=0; i<=3; i++)
		{
			String Strongoi = Getsheet.getRow(0).getCell(i).getStringCellValue();
			System.out.print(Strongoi+" ");
		}
	}

}
