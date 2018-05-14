package PublishGroup;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import PublishGroup.AccessExcel;
import PublishGroup.ExcelRead;
import jxl.read.biff.BiffException;


public class MasterRun 
{
	private static final String req_files = new File(System.getProperty("user.dir")).getParent() + "\\req\\";

	static WebDriver driver;
	static Properties property = new Properties();
	static readingZPL_File read = new readingZPL_File();
	static ArrayList<Map<String, String>>  dlist = new ArrayList<Map<String, String>>();

	public static void main(String[] args) throws IOException, BiffException 
	{
		Logger logger=Logger.getLogger("AirwatchRun");
		PropertyConfigurator.configure(req_files+"log4j.properties"); 

		MasterRun run = new MasterRun();
		try 
		{
			run.runningSheetBySheet(driver, logger);
		} 
		catch (IOException e) 
		{
			read.createErrorFile(driver, logger);
			logger.info("An exception has occured in runtest/copyTCDetails method");
			System.out.println("Failed to load the properties file");
			System.out.println("Exception occured --> "+e.toString());
			driver.quit();
			System.exit(0);
		}
	}


	public void runTest(WebDriver driver, ArrayList<Map<String, String>> dlist,String getGroupname, Logger logger) throws InterruptedException, IOException 
	{
		AccessExcel accessexcel = new AccessExcel();
		readingZPL_File read = new readingZPL_File();
		read.readTCDetailsFile(driver, logger);

		WebElement element;
		Map<String, String> objects;
		
		try
		{
			for (int size = 0; size < dlist.size(); size++) 
			{
				objects = dlist.get(size);
				element = accessexcel.getElement(driver, objects, logger);
				accessexcel.doAction( driver,element, objects,getGroupname,logger);
			}
		}
		catch(Exception e)
		{
			read.createErrorFile(driver, logger);
			logger.error("An Exception has occured in runTest method --> "+e.toString());
			driver.quit();
			System.exit(0);
		}
	}

	public static WebDriver launchBrowser(Logger logger) 
	{
		try
		{
			System.setProperty("webdriver.chrome.driver", req_files+"chromedriver.exe");
			ChromeOptions chromeOptions = new ChromeOptions();
			chromeOptions.addArguments("--start-maximized");
			driver = new ChromeDriver(chromeOptions);
		}
		catch(Exception e)
		{
			read.createErrorFile(driver, logger);
			logger.error("An Exception has occured in launchBrowser method --> "+e.toString());
			driver.quit();
			System.exit(0);
		}
		return driver;
	}


	public void runningSheetBySheet(WebDriver driver,Logger logger) throws IOException, BiffException
	{
		ExcelRead elementAccess = new ExcelRead();
		elementAccess.linkDataSheet(driver, logger);
		elementAccess.publishGroupDataSheet(driver, logger);

		dlist = new ArrayList<Map<String, String>>();
		try
		{	
			String getgroupname= null;
			for (Map<String, String> data : elementAccess.linklist) 
			{
				driver = launchBrowser(logger);
				driver.get(data.get("URL"));

				LinkedHashMap<String, String> productmap = elementAccess.productGroupNameMap(driver, logger);
				ArrayList<String> pd_gp_list = new ArrayList<String>(productmap.values());

				getgroupname = pd_gp_list.get(1);
				
				dlist = elementAccess.publishGrouplist;
				runTest(driver, dlist,getgroupname, logger);
			}
			driver = exitBrowser(logger);
		}
		catch(Exception e)
		{
			read.createErrorFile(driver, logger);
			logger.info("Error in runningSheetBySheet Method "+e.toString());
			driver.quit();
			System.exit(0);
		}
	}


	public static WebDriver exitBrowser(Logger logger)
	{
		logger.info("Successfully Completed the Publishing ZPL Process...");
		driver.quit();
		System.exit(0);
		return driver;
	}
}