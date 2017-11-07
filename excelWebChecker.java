import java.io.File;
import java.io.IOException;
import java.util.Date; 
import jxl.*;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import jxl.read.biff.BiffException; 
import java.util.Scanner;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelWebChecker {
	public static WebDriver driver;
	public static String driverPath = "F:/lib/chromedriver/";

	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException, ParseException {
		System.out.print("Enter path for existing excel spreadsheet. Include .xls at the end of the name.\n"
				+ "Example: ~/Users/aaron/Desktop/newfile.xls with no spaces. \nName: ");
		String filename = scnr.nextLine();

		FileInputStream fs = new FileInputStream(filename);
		Workbook workbook = Workbook.getWorkbook(fs);
		Sheet sheetOne = wb.getSheet(0);

		int numEntries = sheetOne.getPhysicalNumberOfRows(); // returns num rows in the sheet.
		int colNum = 2; //since column two has all of the links.

		ArrayList<String> siteList = new ArrayList<String>();

		for(int i = 0; i< numEntries; i++)
		{
			System.out.println("------------Starting site in entry " + i + " ------------");

			String cellContent = sheetOne.getCell(i, colNum).getContents();

			String site = cellContent;
			System.out.println("site is: " + site);



			//make sure to update the driver path.
 			
			System.setProperty("webdriver.chrome.driver", driverPath + "chromedriver.exe");
			driver = new ChromeDriver();
			driver.get(site);


			//might have to do a wait here...

			String editedUrl = driver.Url;


			if(!editedUrl.contains("https://"))
			{
				System.out.println("Not secure.");

				//this site is not secure, so do not add it, but print it
				siteList.add(site);
			}
			else
			{
				System.out.println("secure!");

				//site is secure, so append https:// and print it
				siteList.add(editedUrl);
			}
			System.out.println("------------Ending site in entry " + i + " ------------");

		}

		driver.quit();


		//write the siteList back to the excel sheet.

	}

}
