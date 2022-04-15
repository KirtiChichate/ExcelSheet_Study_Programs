package excelIntro;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelProgram1 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
//		EXcel count from zero index
		
		
		//Creat fileInputStream class and creat ohject of that class and save that by one obj name and pass a argument in that as path of the excel sheet\\name of excel sheet\\and extension(xlsx)
		FileInputStream ObjFile= new FileInputStream("C:\\Users\\OWNER\\Desktop\\VELOCITY DATA\\Automation Excel sheet\\IntroExcel.xlsx");

/*		//by using some of the methods of FileInputStream we can fetch the data from excel
		String String00 = WorkbookFactory.create(ObjFile).getSheet("IntroExcel").getRow(0).getCell(0).getStringCellValue();
		//here we take only row 0 and column means cell 0 then it shows the results of 00
		System.out.println(String00);
*/
		
/*		//Now if i am trying to find anothe data from same sheet and i am using that sheet name repeatedly without saving that in one ref variable it shows an Exception i.e EmptyFileException
		
		double String06 = WorkbookFactory.create(ObjFile).getSheet("IntroExcel").getRow(0).getCell(6).getNumericCellValue();
		System.out.println(String06);
		
		//Exception---> org.apache.poi.EmptyFileException
		// to avoid this exception 1st u need to store workbookFactory.create in on ref variable then u can use it multiple times
		*/
		
		
		//How to store everything in one ref variable and what its return type that is shown below
		Workbook FileCreate = WorkbookFactory.create(ObjFile);
		
		Sheet Sheetcreate = FileCreate.getSheet("IntroExcel");
		
		Row GetRow = Sheetcreate.getRow(0);
		
		Cell GetCell = GetRow.getCell(0);
		
		String Valueinstring = GetCell.getStringCellValue();  //it is depend on the excel sheet data if what data we r calling i.e in string format then it will becomes getstringcellvalue and if it is in numeric then it becomes getnumericcellvalue.
	//and this is how again we can print it
		System.out.println(Valueinstring);
	}
}
