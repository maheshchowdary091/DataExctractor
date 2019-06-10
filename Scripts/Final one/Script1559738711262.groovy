import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testcase.TestCaseFactory
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords
import internal.GlobalVariable
import MobileBuiltInKeywords as Mobile
import WSBuiltInKeywords as WS
import WebUiBuiltInKeywords as WebUI
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Open excel sheet for reading
FileInputStream fis = new FileInputStream("E:\\Demo1.xlsx");
XSSFWorkbook workbook = new XSSFWorkbook(fis);
XSSFSheet sheet = workbook.getSheet("Sheet1");
def frn = sheet.getLastRowNum();
def lrn = sheet.getFirstRowNum();



//Open new excel file for writing
FileOutputStream fos =new FileOutputStream("E:\\Changedwrite.xlsx");
workbook.write(fos);
XSSFSheet wsheet = workbook.getSheet("Sheet1");


//Used rows counting
int rowCount = lrn-frn;

Row row;
Cell cell;
def val
def bool1
def teamcount
def membername
def name
def firstname
def lastname
def designation
def count =1;

//Open browser and navigating to crunchbase
WebUI.openBrowser('');
WebUI.navigateToUrl('https://www.crunchbase.com/');

//waiting for the Login button and Logging to the application
import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testcase.TestCaseFactory
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords
import internal.GlobalVariable
import MobileBuiltInKeywords as Mobile
import WSBuiltInKeywords as WS
import WebUiBuiltInKeywords as WebUI
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Open excel sheet for reading
FileInputStream fis = new FileInputStream("E:\\Demo1.xlsx");
XSSFWorkbook workbook = new XSSFWorkbook(fis);
XSSFSheet sheet = workbook.getSheet("Sheet1");
def frn = sheet.getLastRowNum();
def lrn = sheet.getFirstRowNum();



//Open new excel file for writing
FileOutputStream fos =new FileOutputStream("E:\\Changedwrite.xlsx");
workbook.write(fos);
XSSFSheet wsheet = workbook.getSheet("Sheet1");


//Used rows counting
int rowCount = lrn-frn;

Row row;
Cell cell;
def val
def bool1
def teamcount
def membername
def name
def firstname
def lastname
def designation
def count =1;

//Open browser and navigating to crunchbase
WebUI.openBrowser('');
WebUI.navigateToUrl('https://www.crunchbase.com/');

//waiting for the Login button and Logging to the application
WebUI.waitForElementClickable(findTestObject('Object Repository/Loginbutton'), 30);
WebUI.click(findTestObject('Object Repository/Loginbutton'));

WebUI.waitForElementClickable(findTestObject('Object Repository/Emailfield'), 30);
WebUI.sendKeys(findTestObject('Object Repository/Emailfield'),'mtools@qualitlabs.com');

WebUI.waitForElementClickable(findTestObject('Object Repository/Passwordfield'), 30);
WebUI.sendKeys(findTestObject('Object Repository/Passwordfield'),'QtlSales@150');

WebUI.waitForElementClickable(findTestObject('Object Repository/Loginbuttonintheform'), 30);
WebUI.click(findTestObject('Object Repository/Loginbuttonintheform'));


for(int i=1;i<=rowCount;i++)
{

	val = findTestData('data').getValue(1, i);


	//Navigate to the Company's Crunchbase url
	WebUI.navigateToUrl(val);

	//checking whether Current team is present or not
	 bool1= WebUI.verifyElementPresent(findTestObject('Object Repository/Team1'), 20);


	if(bool1)
	{
		//checking for the team count
		teamcount = Integer.parseInt(WebUI.getText(findTestObject('Object Repository/TeamCount')))
		
		for(int j=1;j<teamcount;j++)
		{
			membername = WebUI.getText(findTestObject('Object Repository/Teammembername', ["j":j])).trim();
			firstname = membername.split(" ")[0];
		}

		if(teamcount>=1)
		{
			for(int j=1;j<=teamcount;j++)
			{
				//Getting the team member name
				membername = WebUI.getText(findTestObject('Object Repository/Teammembername', ["j":j]));
				name = membername.split(" ");
				firstname = name[0];
				lastname = name[1];

				//Getting the team member designation
				designation = WebUI.getText(findTestObject('Object Repository/Teammemberdesgination', ["j":j]));

				//records filling by reading from input excel sheet
				for(def k = 0;k<10;k++)
				{

					row = wsheet.getRow(count);
					cell = row.createCell(k);
					cell.setCellType(cell.CELL_TYPE_STRING);
					cell.setCellValue(findTestData('data').getValue(k, i));
				}

				//first name, last name and Designation filling in the fields
				cell = row.createCell(10);
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(firstname);

				cell = row.createCell(11);
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(lastname);

				cell = row.createCell(12);
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(designation);

				count=count+1;
			}
		}
	}

	else
	{
		//For no current team -- just paste the record in the excel sheet and place NA in first name, last name and Designation fields

		for(def k = 0;k<10;k++)
		{

			row = wsheet.getRow(count)
			cell = row.createCell(k);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue(findTestData('data').getValue(k, i));


			cell = row.createCell(10);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue("NA");

			cell = row.createCell(11);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue("NA");

			cell = row.createCell(12);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue("NA");

			count = count+1 ;
		}

	}

}





fos.close();
t Repository/Loginbutton'), 30);
WebUI.click(findTestObject('Object Repository/Loginbutton'));

WebUI.waitForElementClickable(findTestObject('Object Repository/Emailfield'), 30);
WebUI.sendKeys(findTestObject('Object Repository/Emailfield'),'mtools@qualitlabs.com');

WebUI.waitForElementClickable(findTestObject('Object Repository/Passwordfield'), 30);
WebUI.sendKeys(findTestObject('Object Repository/Passwordfield'),'QtlSales@150');

WebUI.waitForElementClickable(findTestObject('Object Repository/Loginbuttonintheform'), 30);
WebUI.click(findTestObject('Object Repository/Loginbuttonintheform'));


for(int i=1;i<=rowCount;i++)
{

	val = findTestData('data').getValue(2, i);


	//Navigate to the Company's Crunchbase url
	WebUI.navigateToUrl(val);

	//checking whether Current team is present or not
	bool1= WebUI.verifyElementPresent(findTestObject('Object Repository/Team1'), 20);


	if(bool1)
	{
		//checking for the team count
		teamcount = Integer.parseInt(WebUI.getText(findTestObject('Object Repository/TeamCount')))

		if(teamcount>=1)
		{
			for(int j=1;j<=teamcount;j++)
			{
				//Getting the team member name
				membername = WebUI.getText(findTestObject('Object Repository/Teammembername', ["j":j]));
				name = membername.split(" ");
				firstname = name[0];
				lastname = name[1];

				//Getting the team member designation
				designation = WebUI.getText(findTestObject('Object Repository/Teammemberdesgination', ["j":j]));

				//records filling by reading from input excel sheet
				for(def k = 0;k<10;k++)
				{

					row = wsheet.getRow(count);
					cell = row.createCell(k);
					cell.setCellType(cell.CELL_TYPE_STRING);
					cell.setCellValue(findTestData('data').getValue(k, i));
				}

				//first name, last name and Designation filling in the fields
				cell = row.createCell(10);
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(firstname);

				cell = row.createCell(11);
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(lastname);

				cell = row.createCell(12);
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(designation);

				count=count+1;
			}
		}
	}

	else
	{
		//For no current team -- just paste the record in the excel sheet and place NA in first name, last name and Designation fields

		for(def k = 0;k<10;k++)
		{

			row = wsheet.getRow(count)
			cell = row.createCell(k);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue(findTestData('data').getValue(k, i));


			cell = row.createCell(10);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue("NA");

			cell = row.createCell(11);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue("NA");

			cell = row.createCell(12);
			cell.setCellType(cell.CELL_TYPE_STRING);
			cell.setCellValue("NA");

			count = count+1 ;
		}

	}

}





fos.close();
