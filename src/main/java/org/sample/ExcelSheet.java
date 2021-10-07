package org.sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.formula.WorkbookDependentFormula;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import com.sun.org.apache.bcel.internal.generic.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelSheet {
	public void tc() {
		System.out.println("newone");
		

	}
	private void me2() {
		System.out.println("second");

	}
	private void so() {
     System.out.println("nothing");
	}
	
	
	public static void main(String[] args) throws IOException {
		
	
		
		
		File file = new File("C:\\Users\\Lenovo\\Excel\\driver\\Sheet1.xlsx");
		
		Workbook w = new XSSFWorkbook();
		Sheet sheet = w.createSheet("Sheet");
		
		System.out.println("done");
		
		WebDriverManager.chromedriver().setup();
		
		WebDriver driver = new ChromeDriver();
		
		driver.get("http://demo.automationtesting.in/Register.html");
		
		WebElement sk = driver.findElement(By.id("Skills"));
		
         //sk.click();
		
	  org.openqa.selenium.support.ui.Select select = new org.openqa.selenium.support.ui.Select(sk);
	  
	  List<WebElement> options = select.getOptions();
	  int size = options.size();
	  String text2 = sk.getText();
	  System.out.println(text2);
	 
	 
	  
	  
	 // System.out.println(size);
	  
	  for (int i = 0; i < size; i++) {
		  
		  Row r1 = sheet.createRow(i);
		  Cell cell = r1.createCell(0);
		  WebElement webElement = options.get(i);
		  String text = webElement.getText();
		  cell.setCellValue(text);
		  
	  }	  
		  
		  FileOutputStream str = new FileOutputStream(file);
		  
		  w.write(str);
		  
		  //System.out.println("done");
		  
		  
		  
		
	
	  
	  
		
		
		
		
		
		
		
  
   
   
			
			
			
	}
}
		
		
		
		
	
	
	
	
	
	
	
		
		
		
	
		
		
	      
	     
	      
	      
	      
	    
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	


