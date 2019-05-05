package com.selenium;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

@Test
public class GoogleSearch2 {
    public void test() throws InterruptedException {
    	
    	// Create workbook
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFCreationHelper createHelper = wb.getCreationHelper(); 
        
        HSSFSheet sheet = wb.createSheet("allMickey");
		  // Set header
	  	Row row = sheet.createRow(0);
			row.createCell(0).setCellValue("site name");	
			row.createCell(1).setCellValue("opened url");
			row.createCell(2).setCellValue("number of mickey occurrences");

        // Create a new instance of the Chrome driver
        WebDriver driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(300, TimeUnit.SECONDS);
        driver.manage().window().maximize();

        // Open Chrome with google.com page
        driver.get("http://www.google.com");

        // Find the text input element by it's name
        WebElement element = driver.findElement(By.name("q")); // "q" is name of searching bar
        element.sendKeys("Mickey"); // Search word "Mickey"...it's my nickname ;) 

        // Now submit the form. WebDriver will find the form for us from the element
        element.submit(); // Equal of "enter" button press

        // Google's search is rendered dynamically with JavaScript.
        // Wait for the page to load, timeout after 10 seconds
        (new WebDriverWait(driver, 10)).until(new ExpectedCondition<Boolean>() {
            public Boolean apply(WebDriver d) {
            return d.getTitle().toLowerCase().startsWith("mickey");
            }
        });
        
        ArrayList<String> linksMickey = new ArrayList<String>();
        
        for (int k = 1; k < 11; k++) { 
        			// Find site names that contains 'ickey or 'isney,
        			//Some of the site names do not contain mickey, but they have disney/disneyland      
        			List<WebElement> links = driver.findElements
        			(By.xpath("//div[h2[not(contains(text(),'People also ask'))]]//div[@class='r']//a[contains(.,'ickey')] | //div[h2[not(contains(text(),'People also ask'))]]//div[@class='r']//a[contains(.,'isney')]"));
        		   			// Add links text to array, and write them to console
        					for(WebElement el:links){
        					linksMickey.add(el.getAttribute("href"));	
         					System.out.println(el.getAttribute("href"));
        					}
        					// Click on "next" button from google search
        		driver.findElement(By.xpath("//a[@id='pnnext']")).click();
        		Thread.sleep(2000); 
        }//for
        		// Iterate links from array
        		for(int i = 0; i < linksMickey.size(); i++){
        		driver.get(linksMickey.get(i));
        		// Create row and get text for every cell (page title, url, count "mickey")
        		Row row1 = sheet.createRow(i+1);  
        			row1.createCell(0).setCellValue(createHelper
        				.createRichTextString(driver.getTitle()));
        			row1.createCell(1).setCellValue(createHelper
        				.createRichTextString(driver.getCurrentUrl()));
        			int count = StringUtils.countMatches(driver.getPageSource().toLowerCase(),"mickey");
        			String countMickey = String.valueOf(count);
        			row1.createCell(2).setCellValue(createHelper
        				.createRichTextString(countMickey));
        				// Write to console same info 
        				System.out.println("Page title " + driver.getTitle());
        				System.out.println("Page url " + driver.getCurrentUrl());
        				System.out.println("Number of \"mickey\": " + StringUtils.countMatches(driver.getPageSource().toLowerCase(),"mickey"));
        		}//for
        // Write values to excel file
        try (OutputStream fileOut = new FileOutputStream("mickey2.xls")) {
	         wb.write(fileOut);
	         wb.close();
	         System.out.println("\nFile changed");
	      
	     	} catch (FileNotFoundException e) {
				e.printStackTrace();
	     	} catch (IOException e1) {
				e1.printStackTrace();
			}
     //Close the browser
     driver.quit();
    }//test
}//class
