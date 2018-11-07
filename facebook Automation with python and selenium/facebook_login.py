"""
Target Automation Script
Written by Anthony Silvestri (AS Advanced Analytics)

This script automates the process of downloading sales files from Target site.
Files are downloaded from Target and uploaded to FTP site for processing.
"""

from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import datetime
import time
import csv
import os
import logging
import glob
from retailershared import ftp_upload
from retailershared import retailer_login
from retailershared import list_download


def like_action(driver, like):
    y = 500
    like_id= '//*/a[contains(@class, "UFILikeLink _4x9- _4x9_ _48-k")]'
    for like_button in driver.find_elements_by_xpath(like_id):
		logging.info('enetered in like loop '+str(like))
		driver.execute_script("arguments[0].click();", like_button)
		#driver.find_element_by_xpath(path1).click() #like button
		time.sleep(10)
   		logging.info('liked '+str(like)+'picture')
		like +=1	
		#JavascriptExecutor jse = (JavascriptExecutor)driver;
		driver.execute_script('window.scrollTo(0,"'+str(y)+'")')
		logging.info('scrolled bar ')
		#x=y
		y=y+500	
    return like	

def upload_pic(driver, path_pic):
    try:
	#path_photo = '//*/span/em[contains(text(), "Create a Post")]'
	path_photo = '//*/span[contains(text(), "Create a Post")]'
	for upload_button in driver.find_elements_by_xpath(path_photo):
		logging.info('enetered in upload_pic loop ')
		driver.execute_script("arguments[0].click();", upload_button)
	#WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, path_photo)))
	#logging.info('loacted photo link')
	#WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, path_photo)))
	#driver.find_element_by_xpath(path_photo).click() #photo upload link
	#logging.info('clicked photo upload link ')
	#path_photo1 = '//*/div/input[contains(@title, "Choose a file to upload") and contains(@type, "file")]'
	#path_photo1 = '//*/div[contains(@class, "_5xmp")]/div//div[2]/input'
	path_photo1 = '//*/table[contains(@class, "uiGrid _51mz _5f0n")]//*/div[contains(@class, "_m _5g_r _6a")]/a/div[2]/input'
	for upload_button1 in driver.find_elements_by_xpath(path_photo1):
		logging.info('enetered in upload_pic loop 2 ')
		driver.execute_script("arguments[0].click();", upload_button1)
	time.sleep(30)
	WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, path_photo1)))
	WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, path_photo1)))
	driver.find_element_by_id(path_photo1).click() #photo upload sub link
	logging.info('clicked upload sub link ')


    except Exception, e:
	logging.info(str(e))
	logging.info('error in uploading picture')


def like(driver):
    try:	
	like = 1
	x = 1
	bound = range(1,100)
	try:
	  like_count = like_action(driver, like)
	  for x in bound:
	     if like_count < 10:
		driver.execute_script('window.scrollTo(0,"'+str(600)+'")')
		like_count = like_action(driver, like_count+1)	   
	     else:
	   	logging.info('given the no of likes provided')
		break
	except Exception, e:
	    logging.info(str(e))
	    logging.info('No like button found or sent all likes ')		
    except Exception, e:
	logging.info(str(e))
	logging.info('No like button found ')

def birthday_wish(driver):
    try:	
	path = '//*[@class="fbRemindersStory"]/a/div/div/span'
	WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, path)))
	WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, path)))
	driver.find_element_by_xpath(path).click() #birthday link
	logging.info('clicked birthday link ')
	time.sleep(10)
	#driver.switchTo().alert()
	#element_user = []
	#element_user = glob.glob('u_*_g')
	#lenght_files = len(element_user)
	count = 1
	path2= '//*/textarea[contains(@class, "enter_submit uiTextareaNoResize uiTextareaAutogrow _2yv2 uiStreamInlineTextarea inlineReplyTextArea mentionsTextarea textInput")]'
	try:
	   #no = range(1,3)
	   for input_area in driver.find_elements_by_xpath(path2):
		logging.info('enetered in loop '+str(count))
		driver.execute_script("arguments[0].removeAttribute('onclick');", input_area)
		#path1= '//*[@class="innerWrap"]/textarea'
		#path1= '//*/textarea[contains(@title,"Write on")]'
		#path1= '//*/textarea[contains(@class, "enter_submit uiTextareaNoResize uiTextareaAutogrow _2yv2 uiStreamInlineTextarea inlineReplyTextArea mentionsTextarea textInput")]'
		#input_area = driver.find_element_by_xpath(path1)
		#WebDriverWait(driver, 1800).until(EC.presence_of_element_located((By.XPATH, path1)))
		#WebDriverWait(driver, 1800).until(EC.presence_of_element_located((By.XPATH, path1)))
		#driver.find_element_by_xpath(path1).click() #like button
		time.sleep(10)
    		input_area.send_keys('Happy Bday ')
    		time.sleep(3)
    		input_area.submit()
    		time.sleep(5)
		logging.info('sent bday wish '+str(count))
		count +=1
	   try:
			close_bday_popup = '//*/a[contains(@class, "_42ft _5upp _50zy layerCancel _51-t _50-0 _50z-") and contains(@title, "Close")]'
			WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, close_bday_popup)))
			WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, close_bday_popup)))
			driver.find_element_by_xpath(close_bday_popup).click() #birthday popup close
	   except Exception,e:
			driver.switchTo.alert()
			alert.close()		
	except Exception, e:
	    logging.info(str(e))
	    logging.info('No birthday found or sent birthdays to all ')
	    close_bday_popup = '//*/a[contains(@class, "_42ft _5upp _50zy layerCancel _51-t _50-0 _50z-") and contains(@title, "Close")]'
	    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, close_bday_popup)))
	    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, close_bday_popup)))
	    driver.find_element_by_xpath(close_bday_popup).click() #birthday popup close
	    #driver.switchTo.alert()
	    #alert.close()		
    except Exception, e:
	logging.info(str(e))
	logging.info('No birthday found ')

def logout(driver):
    time.sleep(10)
    try:
	element_logout = '//*/div[contains(text(), "Account Settings")]'
	driver.find_element_by_xpath(element_logout).click()# click on logout list
	#driver.execute_script("arguments[0].click();", element_logout)
	#webdriver.ActionChains(driver).move_to_element(element_logout).click(element_logout).perform()
	logging.info('clicked on logout drop down list ')
	time.sleep(10)
	try:
	    logout_button = '//*/span[contains(@class, "_54nh") and contains(text(), "Log Out")]'
	    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, logout_button)))
	    logging.info('located logout button')
	    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, logout_button)))
	    driver.find_element_by_xpath(logout_button).click() #logout button
	    logging.info('clicked on logout')
	    time.sleep(10)
	except Exception, e:
	    logging.info(str(e))
	    logging.info('exception in clicking logout 1')
	    try:
		logout_button = '//*/span[contains(@class, "_54nh")][8]'
	    	WebDriverWait(driver, 180).until(EC.presence_of_element_located((By.XPATH, logout_button)))
	    	logging.info('located logout button')
	    	WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, logout_button)))
	    	driver.find_element_by_xpath(logout_button).click() #logout button
	    	logging.info('clicked on logout')
	    	time.sleep(10)
	    except Exception, e:
	    	logging.info(str(e))
	    	logging.info('exception in clicking logout 2')
    except Exception, e:
	logging.info(str(e))
	logging.info('exception in logout')



def target_main(target_url, target_user, target_pass, date_format, 
                ftp_server, ftp_user, ftp_pass, ftp_dir, like_count, path_pic):
    """Calls automation functions to download Target sales files
    
    Args:
        target_url: String containing URL for Target site
        target_user: String containing username for Target site
        target_pass: String containing password for Target site
        date_format: Date containing Saturday date ending report week
        ftp_server: String containing the address of the FTP server
        ftp_user: String containing the FTP server username
        ftp_pass: String containing the FTP server password
        ftp_dir: String containing the FTP server directory
    """
    logging.basicConfig(filename='log/facebook_login.log', 
                        format='%(asctime)s : %(levelname)s : %(message)s', 
                        level=logging.INFO)
    try:
        logging.info('Starting to download Target POS files for week ending on ' + str(date_format))
        print 'Starting to download Target POS files for week ending on ' + str(date_format)
        driver = webdriver.Chrome(port=5371)
        driver.get(target_url)
        element_user = 'email'
        element_pass ='pass'
        time.sleep(3)
	retailer_login(driver, target_user, target_pass, element_user, element_pass)
	#no = range(1,11)
	#for count in no:
		#WebDriverWait(driver, 1800).until(EC.presence_of_element_located((By.XPATH, '//*[@class="_khz"][count]/a')))
		#WebDriverWait(driver, 1800).until(EC.presence_of_element_located((By.XPATH, '//*[@class="UFILikeLink _4x9- _4x9_ _48-k"][count]')))
		#path = '//*/form[count]/div/div/div/div/div/div/span/div/a'
		#path = '//*[@class="_khz"]/a'
		#WebDriverWait(driver, 1800).until(EC.presence_of_element_located((By.XPATH, path)))
        	#""" newly added """
        	#WebDriverWait(driver, 1800).until(EC.element_to_be_clickable((By.XPATH, path)))
        	#driver.find_element_by_xpath(path).click() #like button
		#logging.info('clicked like button '+str(count))
	like(driver)
	time.sleep(10)
	driver.execute_script('window.scrollTo(0,0)')
	time.sleep(5)
	birthday_wish(driver)
	logging.info('sent all birthdays')
	logging.info('entering to like function')
	#like(driver)
	logging.info('given no of likes provided')
	#upload_pic(driver, path_pic)
	logout(driver)
	driver.close()
    except Exception, e:
	logout(driver)
        print 'An error has occurred. Check the log file for more information.'
        logging.critical('Error in facebook process', exc_info=True)
        return False

if __name__ == "__main__":
    """Code to run if file is called as script"""
    with open('variables/facebook.csv', 'rb') as csv_file:
        file_dict = csv.DictReader(csv_file)
        for row in file_dict:
            vars()[row['var_name']] = row['value']
    if not date_str:
        date_today = datetime.datetime.today()
        date_format = date_today - datetime.timedelta(days=4)
    else:
        date_format = datetime.datetime.strptime(date_str, '%m/%d/%Y')
    
    target_main(target_url, target_user, target_pass, date_format, 
                ftp_server, ftp_user, ftp_pass, ftp_dir, like_count, path_pic)