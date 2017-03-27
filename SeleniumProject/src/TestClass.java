
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class TestClass {

	/**
	 * @param args
	 * @throws Exception 
	 */
	public static void main(String[] args) throws Exception {
		//int openWindows;
		//System.setProperty("webdriver.chrome.driver", "D:\\MySelWorkSpace\\chromedriver.exe");
		//String processName = "chromedriver.exe";
		
		
		//WebDriver chromeBrowser1 = new ChromeDriver();
		
		//openWindows = chromeBrowser1.getWindowHandles().size();
		//System.out.println(openWindows);
		//if(openWindows>0){
			//chromeBrowser1.quit();
	/*	WebDriver chromeBrowser = new ChromeDriver();	
		chromeBrowser.manage().window().maximize();
		chromeBrowser.navigate().to("https://www.google.co.in");
		System.out.println(chromeBrowser.getTitle());
		chromeBrowser.manage().timeouts().implicitlyWait(5,TimeUnit.SECONDS);
		chromeBrowser.quit();*/
		//}
		
		//System.out.println(file.getCanonicalPath());
        //System.out.println(file.getAbsolutePath());
        //System.out.println(file.getPath());
 //System.getProperty("user.dir"+"\\test.txt");
        
        
		//String strHomePath = System.getProperty("user.home");
		//System.out.println(strHomePath);
		//File fs = new File("\\src\\com\\osf\\testdata\\TestData.xlsx");			
		
		//System.out.println(ft);
		//FileInputStream fin = new FileInputStream(ft);
		
	/*	File file = new File("./src\\TestData.xlsx");        
        FileInputStream ft = new FileInputStream(file);      
		Workbook myWorkbook=null;
		myWorkbook = new XSSFWorkbook(ft);
		Sheet mySheet = myWorkbook.getSheet("TestFile");
		int rowCount = mySheet.getLastRowNum()-mySheet.getFirstRowNum();
		
		Iterator<Row> iRow = mySheet.iterator();
		while(iRow.hasNext()){
			Row row = iRow.next();
			Iterator<Cell> iCell = row.cellIterator(); 
				while(iCell.hasNext()){
					Cell cell = iCell.next();
					
					switch(cell.getCellType()){
					case Cell.CELL_TYPE_STRING:
						System.out.println(cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:	
						System.out.println(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println(cell.getNumericCellValue());
						break;
					}
				}
			System.out.println("");
		}
		ft.close();
	*/
		xlTSRead("./src\\TestData.xlsx");
		
		
		
	}
	public static void xlTSRead(String sPath) throws Exception{
		File myxl = new File(sPath);
		FileInputStream myStream = new FileInputStream(myxl);
		
		XSSFWorkbook myWB = new XSSFWorkbook(myStream);
		XSSFSheet mySheet = myWB.getSheetAt(1);	// Referring to 1st sheet
		int xTSRows = mySheet.getLastRowNum()+1;
		int xTSCols = mySheet.getRow(0).getLastCellNum();
		System.out.println("Rows are " + xTSRows);
		System.out.println("Cols are " + xTSCols);
		String[][] xTSdata = new String[xTSRows][xTSCols];
        for (int i = 0; i < xTSRows; i++) {
	           XSSFRow row = mySheet.getRow(i);
	            for (int j = 0; j < xTSCols; j++) {
	               XSSFCell cell = row.getCell(j); // To read value from each col in each row
	               String value = cellToString(cell);
	               xTSdata[i][j] = value;
	               System.out.println(xTSdata[i][j]);
	               System.out.println("");
	               }
	        }	
	}
	
	public static String cellToString(XSSFCell cell) {
		// This function will convert an object of type excel cell to a string value
	        int type = cell.getCellType();
	        Object result;
	        switch (type) {
	            case XSSFCell.CELL_TYPE_NUMERIC: //0
	                result = cell.getNumericCellValue();
	                break;
	            case XSSFCell.CELL_TYPE_STRING: //1
	                result = cell.getStringCellValue();
	                break;
	            case XSSFCell.CELL_TYPE_FORMULA: //2
	                throw new RuntimeException("We can't evaluate formulas in Java");
	            case XSSFCell.CELL_TYPE_BLANK: //3
	                result = "-";
	                break;
	            case XSSFCell.CELL_TYPE_BOOLEAN: //4
	                result = cell.getBooleanCellValue();
	                break;
	            case XSSFCell.CELL_TYPE_ERROR: //5
	                throw new RuntimeException ("This cell has an error");
	            default:
	                throw new RuntimeException("We don't support this cell type: " + type);
	        }
	        return result.toString();
	    }

}
