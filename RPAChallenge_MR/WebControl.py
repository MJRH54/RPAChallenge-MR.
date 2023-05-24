from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
from os import getcwd
import time


class WebControl:

    def __init__(self):
        self.imageFolderPath = getcwd()+'/images/'
        self.fileSys = FileSystem()
        

    def wait_for_element(self, driver: Selenium ,locator: str):
        """
        Wait for element in the browser.
        """
        maxTime = 0
        timeConf = 0.5
        time_limit = 10
        waitElement = driver.is_element_visible(locator)
        while not waitElement and maxTime < time_limit:
            time.sleep(timeConf)
            maxTime += timeConf
            waitElement = driver.is_element_visible(locator)
        
        if maxTime >= time_limit:
            exit(f'The element {locator} is not present in the browser, execution finished.')
    
    def wait_for_image(self, driver: Selenium ,locator: str):
        """
        Wait for image in the browser.
        """
        time.sleep(0.0001)
        IMAGE_BASE = '//*[@id="site-content"]/div/div[2]/div/ol/'+locator
        if not driver.is_element_visible(IMAGE_BASE):
            for i in range(1,3):
                IMAGE_BASE = f'//*[@id="site-content"]/div/div[{i}]/div[2]/ol/' + locator
                waitElement = driver.is_element_visible(locator)
                if waitElement:
                    break

        return IMAGE_BASE      
    
    def clickElement(self, driver: Selenium, locator: str):
        """
        Click on browser element.
        """
        self.wait_for_element(driver=driver,locator=locator)
        driver.click_element(locator=locator)

    def send_text_to_webBrowser(self,driver: Selenium ,locator: str,text:str):
        """
        Send text to input element.
        """
        self.wait_for_element(driver=driver,locator=locator)
        driver.input_text(locator=locator,text=text)
    
    def send_key(self, driver: Selenium, locator:str, key_to_send='RETURN'):
        """
        Send keys to browser
        """
        self.wait_for_element(driver=driver, locator=locator)
        driver.press_keys(locator, key_to_send)
        time.sleep(1)

    def execute_script(self, driver: Selenium, script:str):
        """
        Execute code javascript
        """
        response = driver.execute_javascript(script)
        return response
        
    def download_image(self, driver:Selenium, locator:str, imageName: str):
        """
        download image from browser.
        """
        if not self.fileSys.does_directory_exist(self.imageFolderPath):
            self.fileSys.create_directory(self.imageFolderPath)
        fileName = self.imageFolderPath + imageName
        xpath = self.wait_for_image(driver=driver,locator=locator)
        driver.capture_element_screenshot(locator=xpath,filename=fileName)
