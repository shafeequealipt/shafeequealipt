
import java.io.FileInputStream;
import java.net.PasswordAuthentication;
import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import java.util.Properties;  
import javax.mail.*;  
import javax.mail.internet.*;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.eclipse.jetty.websocket.api.Session;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;


public class Facebook_Automation {
	
	
	static int like_count;
	static String Username;
	static String Password;
	static String[] country= new String[100];
	
	private static void logout(WebDriver driver) throws Exception
	{
		
		Thread.sleep(10000);
		
		try
		{
			String element_logout = "//*/div[contains(text(), 'Account Settings')]";
			
			driver.findElement(By.xpath(element_logout)).click();
			
			Thread.sleep(5000);
			
			try
			{
				String logout_button = "//*/span[contains(@class, '_54nh') and contains(text(), 'Log Out')]";
				
				driver.findElement(By.xpath(logout_button)).click();
				
				Thread.sleep(5000);
			}
			catch(Exception e)
			{
				System.out.println("error in logout 1");
				
				String logout_button = "//*/span[contains(@class, '_54nh')][8]";
				
				driver.findElement(By.xpath(logout_button)).click();
				
				Thread.sleep(5000);
				
				System.out.println("clicked on logout");
				
			}
			
		}
		
		catch(Exception e)
		{
			System.out.println("Exception in logout");
		}
		
	}
	
	private static void Birthday_wish(WebDriver driver) throws Exception
	{
		try
		{
			String path = "//*[@class='fbRemindersStory']/a/div/div/span";
			
			Thread.sleep(5000);
			
			driver.findElement(By.xpath(path)).click();
			
			Thread.sleep(10000);
			
			String path2= "//*/textarea[contains(@class, 'enter_submit uiTextareaNoResize uiTextareaAutogrow _2yv2 uiStreamInlineTextarea inlineReplyTextArea mentionsTextarea textInput')]";
			
			List<WebElement> blink = driver.findElements(By.xpath(path2));
			for(WebElement bl: blink)
			{
				
				JavascriptExecutor jse = (JavascriptExecutor)driver;
				
				jse.executeScript("arguments[0].removeAttribute('onclick');", bl);
				
				Thread.sleep(8000);
				
				bl.sendKeys("Happy Birthday..");
				
				Thread.sleep(3000);
				
				bl.submit();
				
				Thread.sleep(3000);
				
			}
			
			try
			{
			String close_bday_popup = "//*/a[contains(@class, '_42ft _5upp _50zy layerCancel _51-t _50-0 _50z-') and contains(@title, 'Close')]";
			
			driver.findElement(By.xpath(close_bday_popup)).click();
			}
			
			catch(Exception e)
			{
				//driver.switchTo().alert();
				
				Alert alert = driver.switchTo().alert();
				
				alert.dismiss();
			}
		}
		catch(Exception e)
		{
			System.out.println("No birthday found or sent birthdays to all");
			
			try
			{
			String close_bday_popup = "//*/a[contains(@class, '_42ft _5upp _50zy layerCancel _51-t _50-0 _50z-') and contains(@title, 'Close')]";
			
			driver.findElement(By.xpath(close_bday_popup)).click();
			}
			catch(Exception ex)
			{
				System.out.println("No pop up to close");
			}
		}
	}
	
	public static int like_action(WebDriver driver, int like) throws InterruptedException
	{
		
			int y = 500;
			int s = 0;
			String like_id= "//*/a[contains(@class, 'UFILikeLink _4x9- _4x9_ _48-k')]";

				JavascriptExecutor jse = (JavascriptExecutor)driver;
				
				List<WebElement> allLinks = driver.findElements(By.xpath(like_id));
						
				 for (WebElement we: allLinks)
				 { 

					 try
					 {
						 //driver.findElement(By.xpath(like_id));
						 jse.executeScript("arguments[0].click();", we);
						 System.out.println(like+"like given");
						 Thread.sleep(1000);
				
						 like +=1;
							
					 }
					 
				
				
					 catch(Exception e)
					 {
						 jse.executeScript("window.scrollTo('"+s+"','"+y+"')");
					 }
					
				
					
				
					s=y;
					jse.executeScript("window.scrollTo('"+s+"','"+y+"')");
				
					y = y+500;
				 }

		
		return like;
		
	}
	
	public static void Like(WebDriver driver)
	{
				int like = 1;
				int x = 1;
				//int bound = 100;
				try{
				    like_count = like_action(driver, like);
					for(x=1;x<10;x++)
					{
						if(like_count<10)
						{
							JavascriptExecutor jse = (JavascriptExecutor)driver;
							jse.executeScript("window.scrollTo(0,'"+600+"')");
							Thread.sleep(3000);
							like_count = like_action(driver, like_count+1);
						}
						else
						{
							System.out.println("given the no of likes provided");
							JavascriptExecutor jse = (JavascriptExecutor)driver;
							jse.executeScript("window.scrollTo(0,0)");
							break;
						}
					}
				}
				catch(Exception e)
				{
					System.out.println(e);
					System.out.println("No like button found or sent all likes ");
				}
				
				 
	}
	
	
	public static void Login(WebDriver driver) throws InterruptedException
	{
		try
		{
		String element_user = "email";
		String element_pass ="pass";
	    //Thread.sleep(1000);
		driver.findElement(By.id(element_user)).click();
		
		driver.findElement(By.id(element_user)).sendKeys(Username.toString());
		
		Thread.sleep(2000);
		
		driver.findElement(By.id(element_pass)).click();
		
		driver.findElement(By.id(element_pass)).sendKeys(Password.toString());
		
		Thread.sleep(2000);
		
		driver.findElement(By.id("loginbutton")).click();
		
		Thread.sleep(10000);

		}
		
		catch(Exception ex)
		{
			System.out.println(ex);
		}
		
	}
	
	private static void Read_data(WebDriver driver) throws Exception
	{
	
		 String csvFile = "C:\\Python27\\edge_automation_333\\variables\\facebook.csv";
	        String line = "";
	        String cvsSplitBy = ",";

	        try (BufferedReader br = new BufferedReader(new FileReader(csvFile))) {

	            while ((line = br.readLine()) != null) {

	                // use comma as separator
	                country = line.split(cvsSplitBy);
	                
	                if(country[0].equals("target_user"))
	                {
	                	Username = country[1];
	                }
	                
	                if(country[0].equals("target_pass"))
	                {
	                	Password = country[1];
	                }
	                if(country[0].equals("like_count"))
	                {
	                	like_count =Integer.parseInt(country[1]);
	                }

	                System.out.println(country[0] +  country[1] );

	            }

	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	}

	private static void Sentmail()
	{
		// Set up the SMTP server.
		java.util.Properties props = new java.util.Properties();
		props.put("mail.smtp.host", "smtp.gmail.com");
		//props.put("mail.smtp.host", "localhost");
		//props.put("mail.smtp.host", "usdc2spam2.slingmedia.com");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.port", "587");
		//javax.mail.Session session = javax.mail.Session.getDefaultInstance(props, null);
		
		
		javax.mail.Session session = javax.mail.Session.getDefaultInstance(props,new Authenticator() {
            protected javax.mail.PasswordAuthentication getPasswordAuthentication() {
                return new javax.mail.PasswordAuthentication("shafeequealipt","26160622"); // username and the password
            }
        });
		//javax.mail.Session sess = 

		// Construct the message
		String to = "shafeequealipt@gmail.com";
		String from = "shafeequealipt@gmail.com";
		String subject = "Facebook Automation completed";
		Message msg = new MimeMessage(session);
		try 
		{
		    msg.setFrom(new InternetAddress(from));
		    msg.setRecipient(Message.RecipientType.TO, new InternetAddress(to));
		    msg.setSubject(subject);
		    msg.setText("Hi User,\n\nFacebook Automation for the current date is completed\n\n Thank You.");

		    // Send the message.
		    Transport.send(msg);
		    
		    System.out.println("Sent mail");
		} 
		
		catch (MessagingException e) 
		{
			System.out.println("Exception in Senting mail");
			System.out.println(e);
		}
	 }  
	
	public static void main(String[] args) throws Exception {
		

		WebDriver driver = new ChromeDriver();
		
		driver.get("http://www.facebook.com");
		
		Read_data(driver);
		
		Thread.sleep(5000);
		
		Login(driver);
		
		Thread.sleep(5000);
		
		Like(driver);
		
		Thread.sleep(5000);
		
		Birthday_wish(driver);
		
		Thread.sleep(5000);
		
		logout(driver);
		
		Thread.sleep(5000);
		
		driver.close();
		
		Sentmail();
	}
	
	}

	

	