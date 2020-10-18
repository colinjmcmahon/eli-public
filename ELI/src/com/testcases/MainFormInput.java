//*******************************************************************
// ELI Framework November 2019
//
// 
// 
//*******************************************************************

package com.testcases;

import java.util.regex.Pattern;
import java.util.concurrent.TimeUnit;
import org.junit.*;
import static org.junit.Assert.*;
import static org.hamcrest.CoreMatchers.*;
	import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import jxl.*; 											//jar file added in to work with spreadsheets
import java.io.File;  									//added in for reading in from spreadsheet
import java.io.FileInputStream; 						//added in for reading in from spreadsheet
import java.io.FileNotFoundException;
import java.util.Date;
import java.util.Scanner;
import jxl.write.*; 									//add this in for writing to a spreadsheet
import javax.swing.*; 									//add this package for UI stuff


public class MainFormInput
{
  public WebDriver driver;
  private boolean acceptNextAlert = true;
  private StringBuffer verificationErrors = new StringBuffer();
  
  //Additional Strings for ELI
  private String sProjectPath = "C:\\Users\\Colin\\Google Drive\\ELI Project\\Framework\\";  		//Declare the Project Path
  private String sELI_Input_Test_Data_Sheet_Path;
  private String sInputSheet;
  private String sTestCaseID; 	//Declare variable for Test Case ID from test data input sheet
  private String sTemplate;		//Declare variable for the Template ID from the test data input sheet
  private String sELI_Control_Sheet; //Delcare var for the control sheet
  private String sTemplate_header_on_control_sheet; //Declare var for the template name on the control sheet in row 0 (header row)
  private String sScreen_Template; //Declare var for the screen template as picked up from the control sheet
  private String sAction;
  private String sObject;
  private String sDataString;
  private String sScreen_Template_sheet;
  private String sStringname_on_datasheet;
  private String sDataString_Value;
  private String Obj_English;
  private String sPreferredFormat;
  public String sBy_ID;
  public String sBy_Name;
  private String sBy_xpath;
  private String sDefault_Value;
  private String URL_English;
  private String sBase_URL;
  private String sRestOfURL;
  private int sTest_report_line=0;
  private String sExpected_Repayment_Line;
  private String sActual_Repayment_Line;
  private String sCharAtPos;
  private String sScreen_Template_Form;
  private String sScreen_Template_Var;
  private String sFindVar;
  private String sInclude_Action_Indicator;
  private String sBrowser;
  //*********************************************************************************************
  
    
    
  @Before
  public void setUp() throws Exception 
  {    
	//sBrowser = "Firefox";
    sBrowser = "Chrome";
    
    //Select the location of the test data input sheet
    
    JButton open = new JButton();
	JFileChooser fc = new JFileChooser();
	fc.setCurrentDirectory(new java.io.File("C:/Users/Colin/Google Drive/ELI Project"));
	fc.setDialogTitle("Please select test data input sheet");
	fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
	if (fc.showOpenDialog(open)== JFileChooser.APPROVE_OPTION)
	{
		//
	}

	
	
	sELI_Input_Test_Data_Sheet_Path = fc.getSelectedFile().getAbsolutePath(); //Declare the input sheet to the variable
    
	System.out.println(sELI_Input_Test_Data_Sheet_Path);
	
	switch(sBrowser)
	{
	case "Firefox":
		
		//Declare location of Firefox exe, this avoids errors when updating Firefox.
		System.setProperty("webdriver.firefox.bin", "C://Program Files//Mozilla Firefox//firefox.exe");
		driver = new FirefoxDriver();
		//baseUrl = "http://tcalc.timevalue.com/";
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
    
		break;
    
	case "Chrome":
		
		//Declare location of Chrome exe
		
		 System.setProperty("webdriver.chrome.driver", "C://Users//Colin//Google Drive//ELI Project//SeleniumDrivers//chromedriver_win32//chromedriver.exe"); 
		 
		  driver = new ChromeDriver();
		  driver.manage().window().maximize();
		  driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
					    
	    break;
    
	}    
    
  }
  //*********************************************************************************************
  
  @Test
  public void ELI_Engine_Main() throws Exception
  {
	//sInputSheet = "InputSheet_CLEonline.xls";
	//sInputSheet = "InputSheet_Ryanair.xls";
	sInputSheet = "InputSheet_IrishLifeQuote.xls";
	  
	// *****
	//sELI_Input_Test_Data_Sheet_Path = sProjectPath + "DataSheets\\" + sInputSheet;  //Declare the input sheet to the variable
	// *** 
	
	
	FileInputStream fi_data_input =new FileInputStream(sELI_Input_Test_Data_Sheet_Path);  //File for input
	Workbook w_data_input = Workbook.getWorkbook(fi_data_input); //Workbook for input
	Sheet s_data_input = w_data_input.getSheet(0);  //sheet (first sheet = sheet 0) for input
	
	//Test Report details
	
	 Date date = new Date();
	    
	 String fileName = "C:\\Users\\Colin\\Google Drive\\ELI Project\\Framework\\TestReport.xls";
	 WritableWorkbook testreport = Workbook.createWorkbook(new File(fileName));
	 testreport.createSheet("Sheet10", 0);
	 testreport.createSheet("Sheet12", 1);
	 testreport.createSheet("Sheet13", 2);
	 WritableSheet s_testreport = testreport.getSheet(0);
	 s_testreport.addCell (new Label(0,0,"EXECUTION REPORT " + date ));
	
	//START LOOP - - Test Data Input Sheet *****************************************
	for(int col_input_sheet =1;col_input_sheet<1000;col_input_sheet++) //set the loop to expire on the data input sheet when no of cols=1000... see below comment in 'if' statement..
	 {
		System.out.println("Column on Data Input sheet - Col #" + col_input_sheet);	
		
		sTestCaseID = (s_data_input.getCell (col_input_sheet, 0).getContents());
		sTemplate = (s_data_input.getCell (col_input_sheet, 1).getContents());   // sTemplate = 2nd Row of the data sheet ("s")
		
		System.out.println("Test Case ID at Col #" + col_input_sheet + ": " + sTestCaseID);  //Write out to the console/report log the current test case ID
		System.out.println("Template Name at Col #" + col_input_sheet + ": " + sTemplate);	//Write out to the console the current template selected
		
		//now that we have the template we need to go fetch it from the control sheet..
		if (sTemplate.contains("temp_")) //only if it is a template are we looking it up..
		{
			//sELI_Control_Sheet = sProjectPath + "ControlSheet\\ControlTemp.xls";
			sELI_Control_Sheet = sProjectPath + "ControlSheet\\ControlTemp_Forms.xls";
			FileInputStream fi_control_sheet =new FileInputStream(sELI_Control_Sheet);  //File for control sheet
			Workbook w_control_sheet = Workbook.getWorkbook(fi_control_sheet); //Workbook for control sheet
			Sheet s_control_sheet = w_control_sheet.getSheet(0);  //sheet (first sheet = sheet 0) for control sheet
			
			for(int col_control_sheet =0;col_control_sheet<s_control_sheet.getColumns();col_control_sheet++)  //now read in the control sheet until we have found our template
			{
				sTemplate_header_on_control_sheet = (s_control_sheet.getCell (col_control_sheet, 0).getContents()); //grab the template name from the header (row 0) in the control sheet
				if (sTemplate.equals(sTemplate_header_on_control_sheet))  //If the header on the test data sheet matches the one in the control sheet then go for it.
				{
					System.out.println("Match found on ControlSheet col #" + col_control_sheet + " as " + sTemplate + " = " + sTemplate_header_on_control_sheet);
					//Now that we have found the template name, we need to navigate down on through the rows to read in each screen template
					
					for(int row_control_sheet=1;row_control_sheet<s_control_sheet.getRows();row_control_sheet++)
					{
						sScreen_Template = (s_control_sheet.getCell (col_control_sheet, row_control_sheet).getContents());
						System.out.println("Screen template found on row #" + row_control_sheet +  " of control sheet, column " + col_control_sheet + ": " + sScreen_Template);
						
						for(int m=1;m<sScreen_Template.length();m++)
						{
							//System.out.println("length of " + sScreen_Template + " is: " + sScreen_Template.length());
							
							sCharAtPos = sScreen_Template.substring(m,m+1);
							//System.out.println("sCharAtPos at position " + m + " is: " + sCharAtPos);
							if(sCharAtPos.equals("."))
							{
								
								//then take the substring of sScreen_Template as (0,m-1) ...
								sScreen_Template_Form = sScreen_Template.substring(0,m);
								System.out.println("Screen Template - Form: " + sScreen_Template_Form);
								sScreen_Template_Var = sScreen_Template.substring(m+1,sScreen_Template.length());
								System.out.println("Screen Template - Variation: " + sScreen_Template_Var);
							}
						}
						 
						
						//now open the screen template xls and read through the steps
						
						sScreen_Template_sheet = sProjectPath + "screen_Templates\\" + sScreen_Template_Form + ".xls";
		  	  			FileInputStream fi_screen_template =new FileInputStream(sScreen_Template_sheet);
		  	  			Workbook w_screen_template = Workbook.getWorkbook(fi_screen_template);
		  	  			Sheet s_screen_template = w_screen_template.getSheet(0); 
						
						
						//now read through the steps in the screen template spreadsheet itself
						for(int row_Screen_Template_sheet =0;row_Screen_Template_sheet<s_screen_template.getRows();row_Screen_Template_sheet++)
							 	
						{
							sAction = (s_screen_template.getCell (0, row_Screen_Template_sheet).getContents());
							sObject = (s_screen_template.getCell (1, row_Screen_Template_sheet).getContents());
							sDataString = (s_screen_template.getCell (2, row_Screen_Template_sheet).getContents());
							System.out.println("Action: "  + sAction + ", Object: " + sObject + ", DataString " + sDataString);
							
							//Go find the variation column and then read downwards to see if the actions are included.
							sInclude_Action_Indicator = "N";
							for(int col_var=3;col_var<s_screen_template.getColumns();col_var++)
							{
								sFindVar = (s_screen_template.getCell(col_var, 0).getContents());
								if(sFindVar.equals(sScreen_Template_Var))
								{
									sInclude_Action_Indicator = (s_screen_template.getCell(col_var, row_Screen_Template_sheet).getContents());
									System.out.println("Include Action column Y/N?: " + sInclude_Action_Indicator + " at row " + row_Screen_Template_sheet + ", col: " + col_var);
								}
								else
								{
												
									System.out.println("Include Action column Y/N? - No at row "  + row_Screen_Template_sheet + ", col: " + col_var);
									
								}
								
							}
							
							
							//look up the string and pull back the piece of data from the datasheet
							
							 if (sDataString.contains("$")) 
						    	{
						    	
						    	//then go off and find the relevant datastring in the datasheet
							    	for(int string_lookup_row=1;string_lookup_row<s_data_input.getRows();string_lookup_row++)
						    		{
						    			sStringname_on_datasheet = (s_data_input.getCell (1,string_lookup_row).getContents());

						    			if (sStringname_on_datasheet.equals(sDataString))
						    				{
						    			
						    				sDataString_Value = (s_data_input.getCell (col_input_sheet,string_lookup_row).getContents());
						    				System.out.println("Action: "  + sAction + ", Object: " + sObject + ", DataString " + sDataString);
						    				System.out.println("DataString_Value: " + sDataString_Value);
						    				}	    			    				
						    		}
						    	}
						    else
						    	{
						    	sDataString_Value = sDataString;
						    	//just input what we have then
						    	}
							
							//End of lookup datasheet
							 
							// --------------------------------------------------------------------------- 
							 
							//Look up the object map spreadsheet
							
							//FileInputStream fi_map=new FileInputStream   (sProjectPath + "ObjectMap\\ELI_ObjectsMap.xls"); 
							FileInputStream fi_map=new FileInputStream   (sProjectPath + "ObjectMap\\ELI_ObjectsMap_Forms.xls"); 
							Workbook map = Workbook.getWorkbook(fi_map);
					    	Sheet s_object_map = map.getSheet("ObjectsMapping");
					    		
					    	// --------------------------------------------------------------------------- 
					    	  		
					    	for(int map_row=1;map_row<s_object_map.getRows();map_row++)  //start of for loop for objects map
					    	{
					    			Obj_English = (s_object_map.getCell (0,map_row).getContents());

					    			if (Obj_English.equals(sObject))
					    				{
					    				sPreferredFormat = (s_object_map.getCell (1,map_row).getContents());
					    				System.out.println(Obj_English + ", Preferred Format: " + sPreferredFormat);
					    				
					    				sBy_ID = (s_object_map.getCell (2,map_row).getContents());	
					    				System.out.println(Obj_English + "; Object ID (by ID): " + sBy_ID);
					    				
					    				sBy_Name = (s_object_map.getCell (3,map_row).getContents());	
					    				System.out.println(Obj_English + "; Object ID (by ID): " + sBy_ID);
					    				
					    				sBy_xpath = (s_object_map.getCell (4,map_row).getContents());	
					    				System.out.println(Obj_English + "; Object ID (by xpath): " + sBy_xpath);
					    				
				    					sDefault_Value = (s_object_map.getCell (5,map_row).getContents());	
				    					System.out.println(Obj_English + "; Object ID (by DefaultValue): " + sDefault_Value);
					    				
					    				}	
					    	}
					    	map.close();  //end of for loop for objects map 
							
					    	//Carry out the relevant action on the Object - if the action is allowed based on the Y indicator in the column
					    	
					    	if(sInclude_Action_Indicator.equals("N"))
					    	{
					    		System.out.println("**");
					    		System.out.println("** WARNING: Variation '" + sScreen_Template_Var + "' does not exist. Test Case " + sTestCaseID + " skipped");
					    		System.out.println("**");
					    	}
					    	
					    	else if(sInclude_Action_Indicator.equals("Y"))
					    	{
						    	switch (sAction)
						    	{
						    	  
							    	case "Open URL":
							    		
							    		FileInputStream fi_env=new FileInputStream(sProjectPath + "Environment\\ELI_Environments.xls");  //File for input
							    		Workbook env = Workbook.getWorkbook(fi_env);
							    		Sheet s_env = env.getSheet("Environments");
							    		
							    		for(int env_row=1;env_row<s_env.getRows();env_row++)
							    		{
							    			URL_English = (s_env.getCell (0,env_row).getContents());
							    			
							    			if (URL_English.equals(sObject))
							    			{
							    				System.out.println("URL identified for: " + URL_English);
							    				System.out.println("Got in here!");
							    				sBase_URL = (s_env.getCell (1,env_row).getContents());
							    				sRestOfURL = (s_env.getCell (2,env_row).getContents());
							    			}
							    			;
							    			System.out.println("URL discounted for: " + URL_English);
							    		}
							    		env.close();
							    		driver.get(sBase_URL + sRestOfURL);
				
							    	   	break;
							    	   	
							    	case "TypeInData":
		   					    			    		
							    		if((sPreferredFormat).equals("ID"))
							    		{
							    			System.out.println("Preferred format for field is by 'ID'");
							    			driver.findElement(By.id(sBy_ID)).clear();
							    			driver.findElement(By.id(sBy_ID)).sendKeys(sDataString_Value);
							    			driver.findElement(By.id(sBy_ID)).sendKeys(Keys.TAB);

							    		}
							    		else 
							    		{
							    			if((sPreferredFormat).equals("Name"))
								    		{
							    				System.out.println("Preferred format for field is by 'Name'");
								    			driver.findElement(By.name(sBy_Name)).clear();
								    			driver.findElement(By.name(sBy_Name)).sendKeys(sDataString_Value);
								    			driver.findElement(By.name(sBy_Name)).sendKeys(Keys.TAB);

								    		}
							    			else
							    			{
							    			System.out.println("Preferred format for field is by 'xpath'");							    			
							    			driver.findElement(By.xpath(sBy_xpath)).click();	  
							    			driver.findElement(By.xpath(sBy_xpath)).clear();		
							    			driver.findElement(By.xpath(sBy_xpath)).sendKeys((Keys.BACK_SPACE) + sDataString_Value + (Keys.ENTER));
							    			}
							    		}
							    									    								    		
							    		break;
							    	
							    	case "Test: VerifyFieldContents":
							    		//to be completed...
							    		
							    		break;
							    		
							    		
							    	case "VerifyLabelText":
							    							    		
							    						    		
							    		sTest_report_line ++;
							    		//to be completed...
							    			
							    		sExpected_Repayment_Line = sDataString_Value;
							    							    			
							    		//the following try/catch block has been removed from outside the main block into this case statement
							    		try 
							    		{
							    			// Verify Results
							    			  	      
							    		  	 System.out.println("By.ID = :" + sBy_ID);
							    			 sActual_Repayment_Line = driver.findElement(By.xpath(sBy_ID)).getText();
							    			 System.out.println("Expected Result: " + sExpected_Repayment_Line);
							    			 System.out.println("Actual Result: " + sActual_Repayment_Line);	     
							    			 assertEquals(sExpected_Repayment_Line, driver.findElement(By.xpath(sBy_ID)).getText());
							    			     
							    			 s_testreport.addCell (new Label(3,sTest_report_line,"Pass: "));
											 s_testreport.addCell (new Label(1,sTest_report_line,sExpected_Repayment_Line));
											 s_testreport.addCell (new Label(2,sTest_report_line, sActual_Repayment_Line ));
							    			  	    
							    		  }
							    		      
							    		  catch (Error e)
							    		  {
							    			  verificationErrors.append(e.toString());
							    			  System.out.println("Fail");
							    			  System.out.println("Value returned on screen:" + sActual_Repayment_Line);
							    			  	    
							    			  s_testreport.addCell (new Label(3,sTest_report_line,"Fail: "));
							    			  s_testreport.addCell (new Label(1,sTest_report_line, sExpected_Repayment_Line));
							    			  s_testreport.addCell (new Label(2,sTest_report_line, sActual_Repayment_Line ));
							    			  	    
							    		  }
							    	      //need to add lots more complexity to this option!!!
							    	      //this line currently compares against the old spreadsheet - need to change					    	      	
								  	      			    	
								  	      break;
								  	      
							    	case "Test: Ensure Table Appears":
							    		//to be completed ...
							    		break;
							    	
							    	case "SelectDropDownValue":
							    		
							    		if((sPreferredFormat).equals("ID"))
							    		{
							    			mSelectDropDownValue_By_ID();
							    		}
							    		else if((sPreferredFormat).equals("xpath"))
							    		{
							    			mSelectDropDownValue_By_xpath();
							    		}
							    		
							    		else
							    		{
							    			mSelectDropDownValue_By_LinkText();
								    	}
							    		
							    		break;
							    		
							    	case "Test: Compare Table data":
							    		//to be completed
							    		break;
							    		
							    	case "ClickButton":
							    		
							    		mClickButton();
							    		
							    		break;
							    		
							    	case "ClickOnLink":
							    		
							    		//temporarily redirecting to ClickButton as it has the same code!!
							    		mClickButton();
							    		
							    		break;
							    		
							    	case "ReadValue":
							    		System.out.println("Read Value at line :" + row_Screen_Template_sheet);
							    		
							    		break;
							    		
							    	case "CompareValue":
		
							    		break;
							    		
							    	case "ClickCheckbox":
							    		
							    		mClickCheckbox();
							    		
							    		break;
							    	
							    	case "Select Radio Button":
							    		
							    		mSelectRadioButton();
							    		
							    	case "WriteToLog":
							    		
							    		//Actions ActionsObject = new Actions();
							    		//ActionsObject.WriteToLog(sDataString_Value);
							    		
							    		break;
							    	
							    	case "ClickWeiter":
							    		
							    		driver.findElement(By.id("buttonBlock3")).click();
							    		
							    	default:
							    		
							    		System.out.println(sAction + " is an unknown command!!!");
							    		
						    	}//end of switch action
						    	
					    	} //end of IF statement for if the indicator=Y to include the action
					    	else
					    	{
					    		System.out.println("unknown column inclusion indicator");
					    	}
						    	
					    } //End of reading through the steps in the screen template spreadsheet itself
					    	
							
					}
					
				}
		
				else
				{
					System.out.println("No match found on Control Sheet col #"  + col_control_sheet + " as " + sTemplate + " <> " + sTemplate_header_on_control_sheet);
				}
			}
			
			w_control_sheet.close();
		}
		System.out.println("datasheet col at end loop: " + col_input_sheet);
	 }	//END LOOP - - Test Data Input Sheet *****************************************
	
	w_data_input.close();
  }
  
  //METHODS FOR EACH ACTION IN E.L.I.  ---------------------------------------------------------------
  
  
  public void mSelectDropDownValue_By_ID() throws Exception
  {
		driver.findElement(By.id(sBy_ID)).click();
		driver.findElement(By.id(sBy_ID)).clear();
	    new Select(driver.findElement(By.id(sBy_ID))).selectByVisibleText(sDataString_Value);
	       
	    driver.findElement(By.id(sBy_ID)).sendKeys(sDataString_Value); 
  }
  
  public void mSelectDropDownValue_By_Name() throws Exception
  {
		driver.findElement(By.name(sBy_Name)).click();
		driver.findElement(By.name(sBy_Name)).clear();
	    new Select(driver.findElement(By.name(sBy_Name))).selectByVisibleText(sDataString_Value);
	       
	    driver.findElement(By.name(sBy_Name)).sendKeys(sDataString_Value); 
  }
  
  public void mSelectDropDownValue_By_xpath() throws Exception
  {
		driver.findElement(By.xpath(sBy_xpath)).click();
		//driver.findElement(By.xpath(sBy_xpath)).clear();
		
		 WebElement DropDown = driver.findElement(By.xpath(sBy_xpath));
		    new WebDriverWait(driver, 15).until(
		            ExpectedConditions.elementToBeClickable(DropDown));
		    DropDown.click();
		
	    new Select(driver.findElement(By.xpath(sBy_xpath))).selectByVisibleText(sDataString_Value);
	       
	    //driver.findElement(By.xpath(sBy_xpath)).sendKeys(sDataString_Value);
  }
  
  public void mSelectDropDownValue_By_LinkText() throws Exception
  {
		driver.findElement(By.linkText(sDefault_Value)).click();
	   
	    WebElement DropDown = driver.findElement(By.xpath(sBy_xpath));
	    new WebDriverWait(driver, 15).until(
	            ExpectedConditions.elementToBeClickable(DropDown));
	    DropDown.click();
	    System.out.println("...about to select 1.... ");
	    new Select(driver.findElement(By.id(sBy_ID))).selectByVisibleText(sDataString_Value);
  
  }
     
  public void mSelectRadioButton()
  {
	  driver.findElement(By.id(sBy_ID)).click();
  }
  
  public void mClickButton()
  {
	  System.out.println("LLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLLL" + sBy_ID);
	  
	  
	  driver.findElement(By.id(sBy_ID)).click();
		//need to add lots more complexity to this option!!!
  }
  
  public void mClickCheckbox()
  {
	//Checkbox_yes_no = DataString;
		System.out.println("Checkbox_yes_no :" + sDataString_Value);
		
		switch(sDataString)
		{
			case "Checkbox_YES":
				System.out.println("were in checkbox yes");
				System.out.println("By_ID" + sBy_ID);
				if ( driver.findElement(By.id(sBy_ID)).isSelected() )  //if already selected - do nothing
				{
			     //do nothing
				System.out.println("Checkbox - do not select");
				
				}
				else 
				{//then if not selected - please select
					System.out.println("Checkbox - DO select");
					driver.findElement(By.id(sBy_ID)).click();
				}
				break;
				
			case "Checkbox_NO":
				//System.out.println("were in checkbox No"); 
	
				
				if ( driver.findElement(By.id(sBy_ID)).isSelected() ) 
					    				    				
				{
					System.out.println("checkbox default IS  selected - and were de-selecting it"); 
					driver.findElement(By.id(sBy_ID)).click();
					    					
				}
				else
				{
					//do nothing - if not selected

				}
				break;
				default:
					
		}
		
  }
  
  //*********************************************************************************************
  @After
  public void tearDown() throws Exception 
  {
    driver.quit();
    String verificationErrorString = verificationErrors.toString();
    if (!"".equals(verificationErrorString))
    {
      fail(verificationErrorString);
    }
  }

  private boolean isElementPresent(By by)
  {
    try
    {
      driver.findElement(by);
      return true;
    } catch (NoSuchElementException e)
    {
      return false;
    }
  }

  private boolean isAlertPresent()
  {
    try 
    {
      driver.switchTo().alert();
      return true;
    } catch (NoAlertPresentException e)
    {
      return false;
    }
  }

  private String closeAlertAndGetItsText()
  {
    try {
      Alert alert = driver.switchTo().alert();
      String alertText = alert.getText();
      if (acceptNextAlert)
      {
        alert.accept();
      } else 
      {
        alert.dismiss();
      }
      return alertText;
    } 
    finally 
    {
      acceptNextAlert = true;
    }
  }
//*********************************************************************************************
  
 
}
