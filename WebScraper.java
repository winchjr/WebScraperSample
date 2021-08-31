import org.apache.poi.EncryptedDocumentException;
import javax.script.ScriptException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxBinary;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.firefox.ProfilesIni;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class WebScraper {

	//main, fetch and parser are the only methods you need to edit. 
    public static void main(String[] args) throws EncryptedDocumentException, IOException, InterruptedException, ScriptException {
    	
    	//setting code variables
    	int rowCount = 0;
        int edp = 2090; 
        int arrayCount;
        String header;
        int column = 1;      
        int n = 0;
        List<String> targetEdpDesList = new ArrayList<String>();  //target edp description list
            
        //setting poi variables
        String fileName = "my target excel spreadsheet";     //TODO   
        InputStream inputStream = new FileInputStream(fileName);
        Workbook wb = WorkbookFactory.create(inputStream);
        Sheet sheet = wb.getSheet("All Current Products"); //TODO
		Row row = sheet.getRow(rowCount);    		 			
		Cell cell = row.getCell(0);	
		
		//setting up geckodriver
        System.setProperty("webdriver.gecko.driver","/usr/local/bin/geckodriver");
        File pathBinary = new File("/home/owner/firefox/firefox/firefox");
        FirefoxBinary firefoxBinary = new FirefoxBinary(pathBinary);
        DesiredCapabilities desired = DesiredCapabilities.firefox();
        FirefoxProfile profile = new FirefoxProfile(new File("/home/owner/.mozilla/firefox/xl5jdx0n.default"));
        FirefoxOptions options = new FirefoxOptions();

        //options for geckodriver, prevents anonymous profile from being used.
        desired.setCapability(FirefoxOptions.FIREFOX_OPTIONS, options.setBinary(firefoxBinary)); 
        options.addArguments("--profile", "/home/owner/.mozilla/firefox/xl5jdx0n.default");     
        options.setProfile(profile);
        WebDriver driver = new FirefoxDriver(options);
            
        //I use a static value because I do testing a lot with this, but you can use getPhysicalNumberOfRows with poi to do the whole sheet down to the bottom
        while (edp < 2091){ //TODO      change this code to calculate total number of physical rows, then use that + 1 as the limit.
                           	
        	//get the row of the edp, get the correct webpage for the driver from fetch, and use the parser to parse the webpage retrieved via fetch. parser is the real magic, of converting the target page into data.
        	row = sheet.getRow(edp);
        	driver = fetch(sheet, driver, edp);    
        	targetEdpDesList = parser(driver);    
        	
        	System.out.println(edp);
        	System.out.println(targetEdpDesList);
        	
        	arrayCount = 0;
        	
        	//counts through each element in the array, up until the end. if the element is odd we know its an item description; if even, a header.
        	while (arrayCount < targetEdpDesList.size()) {
        		if (arrayCount % 2 == 0 || arrayCount == 0) { //retrieve the header from the targetEdpDesList, and use findHeaderColumn to get hte header Column number using poi, and set the column number. findHeaderColumn also has a built in feature that, if the header cannot be found, it creates the header for us, and returns the header number column either way.
        			header = targetEdpDesList.get(arrayCount);        			
        			column = findHeaderColumn(sheet, header);
        		}
        		else { //item description. notice that we use column below to place the item description in the correct row.
        			row.getCell(column);
        			row = sheet.getRow(edp);
        			cell = row.createCell(column);
        			cell.setCellValue(targetEdpDesList.get(arrayCount));       			
        		}
        		arrayCount++;
        	}
        	
        	//empty the item description, and go up to the next edp
        	targetEdpDesList.clear();
        	edp++;
	        		
        	//filing our results using poi
	        FileOutputStream outputStream = new FileOutputStream(fileName);       	    
	        wb.write(outputStream);	  
	        outputStream.close();
	        //i moved inputStream.close() and wb.close() to the bottom, so that it writes each new piece to the spreadsheet regardless of whether the resource is closed or not. had quite a few bugs, and couldnt do a full run through with no problems, ever. so just write the correct results as they happen, rather than bulk writing hundreds at a time.
	           
        }        
        inputStream.close();
        wb.close();
        driver.quit();
        System.out.println("complete!");
    }
    
    //fetch does what it sounds like: goes and fetches the website. given a webDriver, the apache poi sheet, and the target EDP on the sheet (given row number), we can look at that item on the website.
	private static WebDriver fetch (Sheet sheet, WebDriver driver, int targetEdp) throws InterruptedException, ScriptException, NoAlertPresentException{
    	
		//setting private variables.
    	Row edpRow = sheet.getRow(targetEdp);
    	Cell edpCell = edpRow.getCell(0);
    	Double edp;
    	String stringEdp;    	
    	WebDriverWait wait = new WebDriverWait(driver, 30); 
  
    	//sometimes the cell is a RichTextString, or a numeric on the sheet. this the trailing zeroes and makes it a string. it also correctly gets the cell value, if its just a string.
    	if (edpCell.getCellType().equals(CellType.NUMERIC)) {
    		edp = edpCell.getNumericCellValue();
    		stringEdp = edp.toString();   	
    		stringEdp = stringEdp.substring(0, stringEdp.length()-1);
    		stringEdp = stringEdp.substring(0, stringEdp.length()-1);
    	}
    	else {
    		stringEdp = edpCell.getStringCellValue();
    	}   				    					
    	//get the website. everything below here has to be custom designed for each website.
        driver.get("insert target website here"); //TODO:        
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[3]/div/div[1]/div[3]/div/div/div[2]/div/div/form/input[1]")));
        Thread.sleep(25);
        driver.findElement(By.xpath("/html/body/div[3]/div/div[1]/div[3]/div/div/div[2]/div/div/form/input[1]")).sendKeys(stringEdp);
        driver.findElement(By.xpath("/html/body/div[3]/div/div[1]/div[3]/div/div/div[2]/div/div/form/input[2]")).click();	
        
		return driver;
	}
	
	private static List<String> parser (WebDriver driver){ //parses the web page, returns a list with item descriptions headers.
		
		 List<String> desList  = new ArrayList<String>();  			
		 WebDriverWait wait = new WebDriverWait(driver, 10);  
		 String price = "";
		 String itemDes1 = "";
		 String series = "";
		 
		 //System.out.println(driver.getPageSource()); //TOD0	
		 
		 //getting price/
		//price seems to be broken or not displaying as of 12/29/20
		/* try {
	     	wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[3]/div[1]/div/div[3]/div/div/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div[2]/div[1]/div[1]/div[1]/span")));
	     	price = driver.findElement(By.xpath("/html/body/div[2]/div[3]/div[1]/div/div[3]/div/div/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div/div/div/div[2]/div[1]/div[1]/div[1]/span")).getText();
	     	itemDes1 = driver.findElement(By.xpath("/html/body/div[2]/div[3]/div[1]/div/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/h1")).getText();	
	     	series = itemDes1.substring(0, itemDes1.indexOf("•"));
	     
		 }
		 catch (Exception e) {
			 desList.add("Failure");
			 desList.add("could not get page fully");
		 } */

	     try {
			 wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[3]/div[2]/div[1]/div/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/h1")));
			 itemDes1 = driver.findElement(By.xpath("/html/body/div[3]/div[2]/div[1]/div/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/h1")).getText(); //the item des
			 series = itemDes1.substring(0, itemDes1.indexOf("•"));
		     desList.add("Price");	     	
		     desList.add(price);		     
		     
		 }
		 catch (Exception e) {
			 
			try { //this is for results that return multiple hits. clicks the top hit which should be the correct one. if this fails then the retrieve in general has failed.
				driver.findElement(By.xpath("/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div[1]/div/div[3]/span")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div[2]/div/div/table/tbody/tr/td[1]/a")));
				driver.findElement(By.xpath("/html/body/div[3]/div/div[3]/div/div/div/div/div/div[2]/div/div/div/div[2]/div[3]/div[1]/div[2]/div/div/table/tbody/tr/td[1]/a")).click();
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[3]/div[2]/div[1]/div/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/h1")));
				itemDes1 = driver.findElement(By.xpath("/html/body/div[3]/div[2]/div[1]/div/div[3]/div/div/div/div/div/div/div[3]/div/div/div/div[2]/div/div/h1")).getText(); //the item des
				series = itemDes1.substring(0, itemDes1.indexOf("•"));
			    desList.add("Price");	     	
			    desList.add(price);	
			}
			catch (Exception e2) {
				desList.add("Failure");
				desList.add("could not get page fully");
			}
		 }
	     Scanner scan = new Scanner(driver.getPageSource());
	     String input;	    
	     	     	     
	     desList.add("Item Description");
	     desList.add(itemDes1);
	     	     	    
	     desList.add("Series");
	     desList.add(series);
	     
	     while (scan.hasNext()) {
	    	 input = scan.next();
	    	
	    	 if (input.contains("class=\"table-container\">")) {
	    		 scan.next();
	    		 scan.next();
	    		 scan.next();
	    		 int tagCount = 0;
	    		 String finalInput = "";	  
	    		 Boolean header = true;
	    		 
	    		 while(!input.contains("tbody")) {
	    			 input = scan.next();	    			 
	    			 tagCount = tagCount + tagCounter(input);	    			
	    			 if ((tagCount == 2 || tagCount == 3) && header == true) { //header description. header bool has to be used because sometimes some tags cause empty headers to be placed.
	    				
		    			 if (tagCount == 3) {//this is the end of the header
		    				 finalInput = finalInput + removeHtmlTags(input);
		    				 desList.add(finalInput); 
		    				 finalInput = "";	
		    				 header = false;
		    			 }
		    			 else {		    				
		    				 finalInput = finalInput + removeHtmlTags(input);
		    			 }
	    			 }
	    			 else if (tagCount == 4 || tagCount == 5) { //item description
		    			 if (tagCount == 5) {//this is the end of the description
		    				 finalInput = finalInput + removeHtmlTags(input);
		    				 desList.add(finalInput); 
	    					 tagCount = -1;
	    					 finalInput = "";
	    					 header = true;
		    			 }
		    			 else {		    				
		    				 finalInput = finalInput + removeHtmlTags(input);
		    			 }	    				 
	    			 }
	    			 
	    		 }
	    	 }
	     }
	     
    	 return desList;   	  	
	
	
	}
    	
//just retrieves the headers column number, or creates one if it doesnt exist.
    private static int findHeaderColumn(Sheet sheet, String headerContent) {
		
    	int cellCount = 0;
    	Row row = sheet.getRow(0);    		 			
    	Boolean foundCell = false;
    	int returnCell = 0;
    	Cell cell = row.getCell(cellCount);	
    	int noOfColumns = sheet.getRow(0).getPhysicalNumberOfCells();
    	
    	for (int x = 0; x < noOfColumns; x++) {//read all the header cells in the first row. if the headerContent matches the header, then return the cell index as an int. if it is not contained at all, then make a new header.

    		
            if (cell.getRichStringCellValue().getString().equals(headerContent)) { 
            	foundCell = true;
                returnCell = cell.getColumnIndex();                   
                }
            else {
            	cellCount++;
            }
            cell = row.getCell(cellCount);
    	}
    	
    	//so we return the cell that is found if a header matches, if it does  not match to any existing headers we create new one and return it.
    	if (!foundCell) {
    	row.getCell(cellCount);
    	cell = row.createCell((noOfColumns));
    	cell.setCellValue(headerContent);
    	return cellCount; //Should return new headers location
    	}
    	else {
    		return returnCell;
    	}
    }
    
    private static String removeHtmlTags (String dirtyString) { //quick code to remove html tags such as <td>. anything between brackets gets removed. use with caution. Now featuring partial html tag removal! 
    	String cleanString = "";
    	if (dirtyString.contains(">") && dirtyString.contains("<")) {
    		cleanString = dirtyString.replaceAll("\\<[^>]*>","");
        	if (cleanString.contains(">")){ //sometimes there is input which looks like this xxyyx>description</endtag> this removes all that before the > too
        		cleanString = cleanString.substring(cleanString.indexOf(">")+1, cleanString.length());  		
        	}
        	else if (cleanString.contains("<")) {
        		cleanString = cleanString.substring(0, cleanString.indexOf("<"));  	
        	}
    	}
    	else if (dirtyString.contains(">")){
    		cleanString = dirtyString.substring(dirtyString.indexOf(">")+1, dirtyString.length());  		
    	}
    	else if (dirtyString.contains("<")) {
    		cleanString = dirtyString.substring(0, dirtyString.indexOf("<"));  	
    	}
    	else {
    		cleanString = dirtyString;
    	}
    	return cleanString;
    }
    
    private static int tagCounter (String input) {//tagCount helps us determine if item descriptions are finished
    	int tagCount = 0;
		for (char c : input.toCharArray()) { 
			if (c == '>') {
				tagCount++;
			}						
		}
		return tagCount;
    }
}
