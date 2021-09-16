package com.taniya.pages;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class AddToCartPage extends BasePage{
    public AddToCartPage(WebDriver driver){
        super(driver);
        PageFactory.initElements(driver, this);

    }


    //WebElements
    @FindBy(xpath = "//div[@class='js-image image card-image']")
    WebElement imgGreenGhost;
    @FindBy(xpath = "(//button[@class='btn btn-primary btn-xs hidden-xs btn-add-to-cart js-addToCart'])[3]")
    WebElement elementAddToCart;
    @FindBy(xpath = "//a[@class='btn btn-primary goto-full-cart']")
    WebElement elementViewCart;

    public void addtoCart() throws InterruptedException {
        actionMove(imgGreenGhost);
        javascriptExecutorsClick(elementAddToCart);
        clickOnElement(elementViewCart);


    }


}
