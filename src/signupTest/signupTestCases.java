package signupTest;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.time.Duration;
import java.util.Random;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class signupTestCases {
	WebDriver driver;
	String website = "https://magento.softwaretestingboard.com/customer/account/create/";
	Random rand = new Random();
	Logger LOGGER = Logger.getLogger(signupTestCases.class.getName());

	@BeforeTest
	public void setUp() {
		driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.get(website);
	}

	@Test()
	public void signupTest() throws EncryptedDocumentException, IOException {
		LOGGER.log(Level.INFO, "Starting signupTest");

		String excelPath = "C:\\Eclipse\\signupTest\\SignupTestData.xlsx";
		FileInputStream file = new FileInputStream(excelPath);

		try {
			LOGGER.log(Level.INFO, "Reading Excel data");
			Workbook workbook = WorkbookFactory.create(file);
			Sheet sheet = workbook.getSheet("Sheet1");

			int rowCount = sheet.getLastRowNum();

			WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

//      Iterate over rows skipping the header row
			for (int i = 1; i <= rowCount; i++) {

				Row row = sheet.getRow(i);

//			get string values from the cells
				String firstName = row.getCell(0).getStringCellValue();
				String lastName = row.getCell(1).getStringCellValue();
				String email = row.getCell(2).getStringCellValue();
				String password = row.getCell(3).getStringCellValue();

//			Locate the sign-up form elements and perform the sign-up operation
				WebElement firstNameInput = wait
						.until(ExpectedConditions.visibilityOfElementLocated(By.id("firstname")));
				WebElement lastNameInput = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("lastname")));
				WebElement emailInput = wait
						.until(ExpectedConditions.visibilityOfElementLocated(By.id("email_address")));
				WebElement passwordInput = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("password")));
				WebElement confirmPasswordInput = wait
						.until(ExpectedConditions.visibilityOfElementLocated(By.id("password-confirmation")));
				WebElement signupButton = wait.until(ExpectedConditions
						.visibilityOfElementLocated(By.cssSelector("button[title='Create an Account'] span")));

//			Input data
				firstNameInput.sendKeys(firstName);
				lastNameInput.sendKeys(lastName);
				emailInput.sendKeys(email);
				passwordInput.sendKeys(password);
				confirmPasswordInput.sendKeys(password);

				signupButton.click();

				WebElement signupMessage = wait.until(ExpectedConditions.visibilityOfElementLocated(
						By.cssSelector("div[data-bind='html: $parent.prepareMessageForHtml(message.text)']")));
				String signupMessageText = signupMessage.getText();

//			Validate sign-up and capture screenshot of the result
				takeScreenshot("./screenshots/" + firstName + lastName + "_Signup.png");
				Assert.assertEquals(signupMessageText, "Thank you for registering with Main Website Store.",
						"Check if sign up message appeared");

				WebElement accountMenuButton = driver
						.findElement(By.cssSelector("div[class='panel header'] button[type='button']"));
				accountMenuButton.click();

				WebElement signOutButton = driver
						.findElement(By.cssSelector("div[aria-hidden='false'] li[data-label='or'] a"));
				signOutButton.click();

				WebElement createAccountButton = driver
						.findElement(By.cssSelector("header[class='page-header'] li:nth-child(3) a:nth-child(1)"));
				createAccountButton.click();
			}
		} catch (Exception e) {
			LOGGER.log(Level.SEVERE, "An error occurred", e);
		} finally {
			file.close();
			driver.quit();
			LOGGER.log(Level.INFO, "signupTest completed");
		}
	}

	public void takeScreenshot(String fileName) throws IOException {
		TakesScreenshot screenshot = (TakesScreenshot) driver;
		File screenshotFile = screenshot.getScreenshotAs(OutputType.FILE);
		File destinationFile = new File(fileName);
		Files.copy(screenshotFile.toPath(), destinationFile.toPath());
        LOGGER.log(Level.INFO, "Screenshot captured: " + fileName);
	}
}