import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.kms.katalon.core.webui.driver.DriverFactory
import org.openqa.selenium.WebDriver
import org.openqa.selenium.By


//findTestData('data').getValue(1, 1)


//***********excel sheet methods*********

FileInputStream fis = new FileInputStream("E:\\dde.xlsx");
XSSFWorkbook workbook = new XSSFWorkbook(fis);
XSSFSheet sheet = workbook.getSheet("Sheet1");
//XSSFSheet sheet = workbook.getSheetAt(0);
def frn = sheet.getLastRowNum();
println frn
def lrn = sheet.getFirstRowNum();
println lrn
int numberOfRows =  frn-lrn


//***********excel sheet methods Completed*********

File f = new File("E:\\latest1.csv")

FileWriter fw = new FileWriter(f)

//***************Login to crunchbase*******************************************
//Open browser and navigating to crunchbase
WebUI.openBrowser('https://www.crunchbase.com/');
WebDriver driver = DriverFactory.getWebDriver()
//WebUI.navigateToUrl('https://www.crunchbase.com/');
WebUI.maximizeWindow()
WebUI.delay(5)

//waiting for the Login button and Logging to the application
WebUI.waitForElementClickable(findTestObject('Object Repository/Loginbutton'), 30);
WebUI.click(findTestObject('Object Repository/Loginbutton'));
WebUI.waitForElementClickable(findTestObject('Object Repository/Emailfield'), 30);
WebUI.sendKeys(findTestObject('Object Repository/Emailfield'),'mtools@qualitlabs.com');
WebUI.waitForElementClickable(findTestObject('Object Repository/Passwordfield'), 30);
WebUI.sendKeys(findTestObject('Object Repository/Passwordfield'),'QtlSales@150');
WebUI.waitForElementClickable(findTestObject('Object Repository/Loginbuttonintheform'), 30);
WebUI.click(findTestObject('Object Repository/Loginbuttonintheform'));
//***************Login complete*******************************************

//*******Print Headers in csv**************


def Headers = ["Organization Name", "URL", "Categories", "Headquarters Location", "Description", "CB Rank (Company)", "Last Funding Date", "Website", "Founders", "Founders", "Number of Employees", "First Name", "Last Name", "Designation"]

for(Header in Headers)
{
	fw.append(Header);
	fw.append(',');
}
fw.append(" ")
fw.append('\n')



//*******Completed Print Headers in csv**************


def firstName
def lastName
def designation

//int refCount = 0

def teamcount
for(int i=1;i<=numberOfRows;i++)
{
	println findTestData('data').getValue(2, i)
	WebUI.navigateToUrl(findTestData('data').getValue(2, i))
	WebUI.waitForPageLoad(30)
	def bool = WebUI.verifyElementPresent(findTestObject('Object Repository/Team1'), 10, FailureHandling.OPTIONAL);
	//teamcount = Integer.parseInt(WebUI.getText(findTestObject('Object Repository/TeamCount')))
	for(int k=1;k<=10;k++)
	{
		fw.append(findTestData('data').getValue(k, i))
		fw.append(",")
	}
	if(bool)
	{
		teamcount = driver.findElements(By.xpath("//span[text()='Number of Current Team ']/../../../../..//a")).size()
		//teamcount = Integer.parseInt(WebUI.getText(findTestObject('Object Repository/TeamCount')))
		for(int tm=1;tm<=teamcount;tm++)
		{
			if(tm>1)
			{
				fw.append("\n")
				for(int sp=1;sp<=10;sp++)
				{
					fw.append(" ")
					fw.append(",")
				}
			}
			WebUI.delay(3)
			firstName = (WebUI.getText(findTestObject('Object Repository/Teammembername', ["j":tm])).trim()).split(" ")[0]
			lastName = (WebUI.getText(findTestObject('Object Repository/Teammembername', ["j":tm])).trim()).split(" ")[1]
			designation = WebUI.getText(findTestObject('Object Repository/Teammemberdesgination', ["j":tm]));
			fw.append(firstName)
			fw.append(",")
			fw.append(lastName)
			fw.append(",")
			fw.append(designation)
			fw.append('\n')
		}
	}
	else
	{
		fw.append("NA")
		fw.append(",")
		fw.append("NA")
		fw.append(",")
		fw.append("NA")
		fw.append("\n")
	}
}
fw.flush()
fw.close()




