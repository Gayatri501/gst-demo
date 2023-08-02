package gstdemo;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class launchbrowser {
    public static void main(String[] args) {
        // Set the system property to specify the location of the Firefox driver (geckodriver)
    	WebDriverManager.firefoxdriver().setup();
		WebDriver driver = new FirefoxDriver();
     
        driver.get("https://www.trivago.in/?cip=91030227040912&cip_tc=12891-101-101_privacy_1689937781401000000&cip_ext_id=1689937781401000000&mfadid=adm");

        driver.findElement(By.xpath("//span[normalize-space()='Search']")).click();
      

        // Close the browser
       // driver.quit();
    }
}
