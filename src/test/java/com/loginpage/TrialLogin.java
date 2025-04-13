package com.loginpage;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.*;

public class TrialLogin {

    WebDriver driver;

    @BeforeClass
    public void setUp() {
        System.setProperty("webdriver.chrome.driver", "D://chromedriver135.exe");
        driver = new ChromeDriver();
        driver.get("https://itassetmanagementsoftware.com/rolepermission/admin/login");
        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
    }

    @Test
    public void testLoginWithExcelData() throws Exception {

    	
    	    String path =   System.getProperty("user.dir")+"/src/test/resources/DataLogin.xlsx";
    	    FileInputStream fis = new FileInputStream(path);
    	    Workbook wb = WorkbookFactory.create(fis);
    	    Sheet sheet = wb.getSheet("SHEET3");

    	    DataFormatter formatter = new DataFormatter();
    	    int totalRows = sheet.getPhysicalNumberOfRows();
    	    System.out.println("Total rows (excluding header): " + (totalRows - 1));

    	    for (int i = 1; i < totalRows; i++) {
    	        Row row = sheet.getRow(i);
    	        if (row == null) continue;

    	        Cell usernameCell = row.getCell(0);
    	        String username = formatter.formatCellValue(usernameCell).trim();
    	        System.out.println("Trying username at row " + (i + 1) + ": " + username);

    	        WebElement usernameField = driver.findElement(By.cssSelector("#username"));
    	        usernameField.clear();
    	        usernameField.sendKeys(username);

    	        driver.findElement(By.xpath("//span[text()='Continue']")).click();
    	      
    	        String message = "";
    	        try {
    	            // Attempt to capture error message
    	            WebElement errorElement = driver.findElement(By.xpath("//div[contains(@class, 'error-message') or contains(@class, 'alert')]"));
    	            message = errorElement.getText().trim();
    	        } catch (NoSuchElementException e) {
    	            // If no error message is found, handle this case
    	            try {
    	                JavascriptExecutor js = (JavascriptExecutor) driver;
    	                message = (String) js.executeScript(
    	                    "var el = document.querySelector('#username');" +
    	                    "if (el && el.nextElementSibling) {" +
    	                    "    return el.nextElementSibling.innerText.trim();" +
    	                    "} else {" +
    	                    "    return 'No error message found near username';" +
    	                    "}"
    	                );
    	            } catch (Exception ex) {
    	                message = "Error capturing message";
    	            }
    	        }

    	        // If the username is blank, check for a specific error message related to that
    	        if (username.isEmpty()) {
    	            if (message.isEmpty()) {
    	                message = "Username field cannot be empty";  // Custom message when the field is blank
    	            }
    	        }

    	        System.out.println("Captured message for '" + username + "': " + message);

    	        // Write message to column 1 (B) in Excel if message is not empty
    	        if (message != null && !message.trim().isEmpty()) {
    	            Cell messageCell = row.getCell(1);
    	            if (messageCell == null) {
    	                messageCell = row.createCell(1);
    	            }
    	            messageCell.setCellValue(message);
    	        } else {
    	            System.out.println("No message for username '" + username + "', skipping write to Excel.");
    	        }
    	    }

    	    fis.close();
    	    FileOutputStream fos = new FileOutputStream(path);
    	    wb.write(fos);
    	    fos.close();
    	    wb.close();

    	    driver.quit();
    	    System.out.println(" Test completed. Results written to Excel.");
    	}
}
