package com.ipl.base;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

public class TestBase {
    public WebDriver driver;

    @BeforeClass // Changed from @BeforeMethod to fix NullPointerException
    public void setup() {
        ChromeOptions options = new ChromeOptions();
        
        options.addArguments("--no-sandbox");
        options.addArguments("--disable-dev-shm-usage");
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--start-maximized");
        options.addArguments("--disable-blink-features=AutomationControlled");
        
        
        options.addArguments("user-data-dir=C:/selenium-chrome-profile");;

        driver = new ChromeDriver(options);
    }

    @AfterClass // Changed to @AfterClass to keep browser open for all DataProvider rows
    public void tearDown() {
        if (driver != null) {
            driver.quit();
        }
    }
}