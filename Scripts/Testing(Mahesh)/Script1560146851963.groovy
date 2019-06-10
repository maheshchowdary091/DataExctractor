import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testcase.TestCaseFactory
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository
import com.kms.katalon.core.testobject.TestObject
import internal.GlobalVariable
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


//Open excel sheet for reading
/*
 FileInputStream fis = new FileInputStream("E:\\Demo1.xlsx");
 XSSFWorkbook workbook = new XSSFWorkbook(fis);
 XSSFSheet sheet = workbook.getSheet("Sheet1");
 def frn = sheet.getLastRowNum();
 def lrn = sheet.getFirstRowNum();
 */

File src = new File("E:\\Changedwrite.xlsx")
FileInputStream fis = new FileInputStream(src);
XSSFWorkbook workbook = new XSSFWorkbook(fis);


XSSFSheet sheet = workbook.getSheet("Sheet1");


//Open new excel file for writing
FileOutputStream fos =new FileOutputStream(src);

def Headers = ["Organization Name", "URL", "Categories", "Headquarters Location", "Description", "CB Rank (Company)", "Last Funding Date", "Website", "Founders", "Founders", "Number of Employees", "First Name", "Last Name", "Designation"]
//Used rows counting
//int rowCount = lrn-frn;
row=sheet.createRow(0)

//cell = row.getCell(0)

for(int nof=0;nof<13;nof++)
{
cell = row.createCell(nof);
cell.setCellType(cell.CELL_TYPE_STRING);
cell.setCellValue(Headers[nof]);
}


workbook.write(fos);