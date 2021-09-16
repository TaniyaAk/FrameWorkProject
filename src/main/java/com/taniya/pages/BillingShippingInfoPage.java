package com.taniya.pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class BillingShippingInfoPage extends BasePage {
    public BillingShippingInfoPage(WebDriver driver) {
        super(driver);
        PageFactory.initElements(driver, this);
    }

    //WebElements
    @FindBy(id = "shouldUseShippingAddress1")
    static
    WebElement elementSameShippingAddress;
    @FindBy(xpath = "//a[@href='#CreditCard']")
    static
    WebElement elementCreditCard;
    @FindBy(id = "cardNumber")
    static
    WebElement elementCardNumber;
    @FindBy(id = "securityCode")
    static
    WebElement elementSecurityCode;
    @FindBy(id = "cardExpDate")
    static
    WebElement elementCardExpDate;
    @FindBy(xpath = "(//*[@class='check'])[2]")
    static
    WebElement elementSaveTheCard;
    @FindBy(xpath = "(//i[@class='material-icons'])[7]")
    static
    WebElement elementContinueBtnInput;
    @FindBy(xpath = "(//a[@class='btn btn-primary pull-right js-performCheckout'])[1]")
    static
    WebElement elementPlaceOrder;


    //created method here

    public static void billingShippingInfoPage() throws InterruptedException {
        javascriptExecutorsClick(elementSameShippingAddress);
        clickOnElement(elementCreditCard);
        elementCardNumber.clear();
        textInput(elementCardNumber, "22222222222");
        elementSecurityCode.clear();
        textInput(elementSecurityCode, "302");
        elementCardExpDate.clear();
        textInput(elementCardExpDate, "04/23");
        clickOnElement(elementSaveTheCard);
        clickOnElement(elementContinueBtnInput);
        Thread.sleep(5000);
        clickOnElement(elementPlaceOrder);
    }
}
