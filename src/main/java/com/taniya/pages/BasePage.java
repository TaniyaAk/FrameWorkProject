package com.taniya.pages;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

public class BasePage {
    String email = "taniya4@me.com";
    private WebDriver driver;
    //create a constructor

    public BasePage(WebDriver driver) {
        this.driver = driver;
    }

    protected WebDriver getDriver() {
        return this.driver;
    }

    //This method clicks on an element
    public static void clickOnElement(WebElement element) throws InterruptedException {
        Thread.sleep(2000);
        //highlightElement(element);
        element.click();
        element.getText();
    }

    public void highlightElement(WebElement element) {
        JavascriptExecutor jse = (JavascriptExecutor) driver;
       // jse.executeScript("argument[0].style.border='solid 3px yellow'", element);
    }

    public void actionMove(WebElement element) {
        Actions action = new Actions(driver);
        action.moveToElement(element);
    }

    public static void javascriptExecutorsClick(WebElement element) {
        //JavascriptExecutor js = (JavascriptExecutor) driver;
       // js.executeScript("arguments[0].click();", element);
    }

    //This method enters text in a textBox
    public static void textInput(WebElement element, String textToType) {
        //highlightElement(element);
        element.clear();
        element.sendKeys(textToType);
    }
}
