import os
import time
import tablib
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

class Verint:

    def __init__(self, storeNum, password, saveDir):
        """automate getting data from verint review"""
        self.storeNum = str(storeNum)
        self.fullStoreNum = ''

        if (len(self.storeNum)) == 1:
            self.fullStoreNum = '000' + self.storeNum
        if (len(self.storeNum)) == 2:
            self.fullStoreNum = '00' + self.storeNum
        if (len(self.storeNum)) == 3:
            self.fullStoreNum = '0' + self.storeNum
        else:
            self.fullStoreNum = self.storeNum
            
        self.fullUrl = 'beginning' + self.fullStoreNum + 'end'

        self.password = password
        self.saveDir = saveDir
        self.driver = ''
        self.actions = ''
        self.searchWindow = ''
        self.previewErrorCount = 0
        self.exitStatus = ''

    def startDriver(self):
        """Initialize driver and actions"""
        self.driver = webdriver.Remote(
            command_executor='http://localhost:9999',
            desired_capabilities={
                "debugConnectToRunningApp": 'false',
                "app": r"C:\Program Files (x86)\Verint\Review\Review.exe",
                'launchDelay': '2'
            })  
        self.actions = ActionChains(self.driver)

    def login(self):
        """Handles the login screen"""
        windowFound = False
        while(windowFound == False):
            try:
                time.sleep(3)
                loginWindow = self.driver.find_element_by_id('LoginView')
                url = loginWindow.find_element_by_class_name('Edit')
                url.clear()
                url.send_keys(self.fullUrl)

                password = loginWindow.find_element_by_id('m_Password')
                    
                if password.is_enabled():
                    password.clear()
                    password.send_keys(self.password)

                loginButton = loginWindow.find_element_by_id('m_LoginButton')
                loginButton.click()
                windowFound = True
            except NoSuchElementException:
                continue
        
            

    def openSearch(self):
        """Open the camera search pane"""
        windowFound = False
        while(windowFound == False):
            try:
                self.driver.find_element_by_name('Cameras')
                windowFound = True
            except NoSuchElementException:
                continue
        
        toggleButton = self.driver.find_element_by_name('ngToggleButton1')
        toggleButton.click()

        self.actions.move_by_offset(10, 0).perform()
        self.actions.click().perform()
        time.sleep(3)

        self.searchWindow = self.driver.find_element_by_id('m_TopWorkspace')

        expander = self.searchWindow.find_element_by_name('m_CameraNameExpander')
        expander.click()

    def cameraSearch(self, cameraName):
        """Filter to and open a single camera view"""
        textBox = self.searchWindow.find_element_by_id('m_CameraNameTextBox')
        textBox.clear()
        textBox.send_keys(cameraName)

        searchButton = self.searchWindow.find_element_by_name('m_SearchBtn')
        searchButton.click()
        time.sleep(1)

        pyautogui.doubleClick(x = 50, y = 535)              

    def openImage(self, cameraName):
        """Opens the snapshot tool"""
        windowFound = False
        while(windowFound == False):
            try:
                preview = self.driver.find_element_by_id('NextivaVideoControlPlayBarOverlay')
                time.sleep(2)
                windowFound = True
            except NoSuchElementException:
                continue

        time.sleep(15)
        pyautogui.hotkey('ctrl', 'i')
        time.sleep(2)
        
        windowFound = False
        while(windowFound == False):
            try:
                save = self.driver.find_element_by_id('Item 57603')
                save.click()
                time.sleep(2)
                windowFound = True

            except NoSuchElementException:
                self.previewErrorCount += 1
                print(f'{self.storeNum} Error opening preview: {self.previewErrorCount}')
                
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.hotkey('ctrl', 'i')
                time.sleep(2)

                if self.previewErrorCount > 4:
                    self.exitStatus = 'ERROR'
                    return

    def saveImage(self, cameraName):
        """saves the image file"""
        saveAs = self.driver.find_element_by_name('Save As')
        filePath = saveAs.find_element_by_name('Address: Documents')
        filePath.click()
        filePath.send_keys(self.saveDir)
        filePath.submit()
        time.sleep(1)
            
        filename = saveAs.find_element_by_id('1001')
        fileTimeStamp = time.strftime('%m.%d.%Y %H.%M', time.gmtime(time.time()))
        file = self.storeNum + ' ' + cameraName + ' ' + fileTimeStamp
        filename.send_keys(file)
        time.sleep(1)

        pyautogui.hotkey('enter')
        time.sleep(3)
        pyautogui.hotkey('alt', 'space', 'c')
        return(f'{file}.bmp')

    def snapshot(self, cameraNames):
        """Saves snapshot images for a list of camera"""
        startTime = time.time()
        self.startDriver()
        self.login()
        self.openSearch()

        for cameraName in cameraNames:
            self.cameraSearch(cameraName)
            self.openImage(cameraName)

            if self.previewErrorCount <= 4:
                saveSuccess = False
                while(saveSuccess == False):
                    file = self.saveImage(cameraName)
                    os.chdir(self.saveDir)
                    saveSuccess = os.path.isfile(file)
                        
                    if saveSuccess == False:
                        print(f'{self.storeNum}: {file} was not saved')
                        self.openImage(cameraName)
                    else:
                        print(f'{self.storeNum}: {file} saved ok!')
                        self.previewErrorCount = 0
                        self.exitStatus = 'OK'

        if self.exitStatus != 'ERROR':
            totalTime = time.time() - startTime
            print(f'{self.storeNum}: {totalTime:.2f} seconds for {len(cameraNames)} cameras')
            
        self.driver.close()
        print(f'{self.storeNum} exit status: {self.exitStatus}')
        
        
        

