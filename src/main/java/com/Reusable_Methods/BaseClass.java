package com.Reusable_Methods;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.safari.SafariDriver;


public class BaseClass {

    public WebDriver driver;

    public void launchBrowser(String browser) {
        switch (browser) {
            case "chrome":
                WebDriverManager.chromedriver().setup();
                driver = new ChromeDriver();
                break;

            case "edge":
                WebDriverManager.edgedriver().setup();
                driver = new EdgeDriver();
                break;

            case "firefox":
                WebDriverManager.firefoxdriver().setup();
                driver = new FirefoxDriver();
                break;

            case "IE":
                WebDriverManager.iedriver().setup();
                driver = new InternetExplorerDriver();
                break;

            case "safari":
                WebDriverManager.safaridriver().setup();
                driver = new SafariDriver();
                break;
        }
    }
}
