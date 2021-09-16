package com.taniya.pages;


import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class LoginPage extends BasePage {

    //Create a Constructor
    public LoginPage(WebDriver driver) {
        super(driver);
        PageFactory.initElements(driver, this);

    }

    //WebElements
    @FindBy(xpath = "//a[@href ='/login']")
    WebElement loginBtnTwo;
    @FindBy(xpath = "//input[@id='username']")
    WebElement registerEmailInputTwo;
    @FindBy(xpath = "(//*[@id='password'])[1]")
    WebElement enterPasswordTwo;
    @FindBy(className = "check")
    WebElement rememberMeInput;
    @FindBy(xpath = "(//button[@type='submit'])[2]")
    WebElement loginButtonTwo;


    public void login() throws InterruptedException {
        clickOnElement(loginBtnTwo);
        textInput(registerEmailInputTwo, email);
        textInput(enterPasswordTwo, "PW1234");
        clickOnElement(rememberMeInput);
        clickOnElement(loginButtonTwo);
        Thread.sleep(1000);
    }
    public void login(String userID, String pass) throws InterruptedException{
        clickOnElement(loginBtnTwo);
        textInput(registerEmailInputTwo, userID);
        textInput(enterPasswordTwo, pass);
        clickOnElement(rememberMeInput);
        clickOnElement(loginButtonTwo);
    }
}


