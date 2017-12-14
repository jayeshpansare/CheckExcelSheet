package checkExcelSheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import org.apache.http.util.Asserts;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class checkExcelSheet {

	private static final Object[][] Object = null;
	private static Object name;

	public static void main(String[] args) throws IOException {
		
		/**
		 * get the excel sheet  
		 * 
		 ***/
		File src = new File("C:\\Users\\jayesh\\Desktop\\checkdata.xlsx");
		FileInputStream fis=new FileInputStream(src);
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		wb.getAllNames();
		
		XSSFSheet sh1 = wb.getSheetAt(0);
		int getrowcount = sh1.getLastRowNum();
		
		ArrayList<Object> data = new ArrayList<Object>();
		
		for(int i=0; i<getrowcount; i++){
			
			try{	
				String getname = sh1.getRow(i).getCell(0).getStringCellValue();
				data.add(getname);
			}catch(Exception e){
				
			}			
		}
		System.out.println("Read Excel Sheet");
		
		/**
		 * remove dublicate value 
		 * 
		 ***/
		for(int i=0; i<data.size();i++){
			for(int j=i+1; j<data.size();j++){
			
				if(data.get(i).equals(data.get(j))){
					data.remove(j);
					j--;
				}
			}
		}
		System.out.println("Remove Dublicate Value");
		
		
		/**
		 *  Write File
		 * ***/
		File writefile = new File("C:\\Users\\jayesh\\Desktop\\checkdata_write.xlsx");
		FileInputStream inputStream = new FileInputStream(writefile);
		
		XSSFWorkbook Workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = Workbook.getSheetAt(0);
		for(int j=0; j<data.size();j++){
			Row row = sheet.getRow(j);
			Row newRow = sheet.createRow(j);
			Cell cell = newRow.createCell(0);
			String get_data = data.get(j).toString();
			cell.setCellValue(get_data);
		}
		
		
		inputStream.close();
		FileOutputStream outputStream = new FileOutputStream(writefile);
		Workbook.write(outputStream);
		outputStream.close();
		System.out.println("Write Excel Sheet");
	}
}
