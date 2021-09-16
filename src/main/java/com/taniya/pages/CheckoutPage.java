package com.taniya.pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class CheckoutPage extends BasePage {
    public CheckoutPage(WebDriver driver) {
        super(driver);
        PageFactory.initElements(driver, this);
    }

    //WebElements
    @FindBy(xpath = "//a[@class='btn btn-primary']")
    static
    WebElement elementCheckoutOne;
    @FindBy(xpath = "//*[@id='fullName']")
    static
    WebElement elementFullName;
    @FindBy(id = "addressLine1")
    static
    WebElement elementAddressOne;
    @FindBy(id = "addressLine2")
    static
    WebElement elementAddressTwo;
    @FindBy(id = "city")
    static
    WebElement elementCity;
    @FindBy(id = "stateProvinceRegion")
    static
    WebElement elementState;
    @FindBy(id = "postalCode")
    static
    WebElement elementPostalCode;
    @FindBy(id = "phonePrimary")
    static
    WebElement elementPhoneNumber;
    @FindBy(xpath = "//*[@id='shippingInfo']/div/div[1]/div/div/div[5]/div/div")
    static
    WebElement elementCheckBox;
    @FindBy(id = "fulfillmentOptionId1")
    static
    WebElement elementShippingMethod;
    @FindBy(xpath = "//a[@class='btn btn-primary pull-right js-submitCheckoutStage'] ")
    static
    WebElement elementCheckoutTwo;


    public static void checkoutPage() throws InterruptedException {
        clickOnElement(elementCheckoutOne);
        textInput(elementFullName, "Awesome Taniya");
        textInput(elementAddressOne, "243 Gold St");
        textInput(elementAddressTwo, "13B");
        textInput(elementCity, "Brooklyn");
        textInput(elementState, "NY");
        textInput(elementPostalCode, "11201");
        textInput(elementPhoneNumber, "800-222-2000");
        clickOnElement(elementCheckBox);
        javascriptExecutorsClick(elementShippingMethod);
        clickOnElement(elementCheckoutTwo);
        Thread.sleep(20000);

    }
}
