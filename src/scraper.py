"""
Web scraper module for VTU Result Automation
Handles Selenium operations to scrape student results from the VTU website
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
import logging
from logger import setup_logger
from config import (
    XPATH_USN_INPUT,
    XPATH_CAPTCHA_INPUT,
    XPATH_SUBMIT_BUTTON,
    XPATH_STUDENT_USN,
    XPATH_STUDENT_NAME,
    XPATH_SUBJECT_BASE,
    SUBJECT_NAME_INDEX,
    SUBJECT_IA_INDEX,
    SUBJECT_SEE_INDEX,
    SUBJECT_TOTAL_INDEX,
    SUBJECT_RESULT_INDEX,
    MAX_SUBJECTS,
    WAIT_AFTER_STARTUP,
    WAIT_BEFORE_CAPTCHA,
    WAIT_AFTER_INPUT,
    WAIT_TIMEOUT
)

logger = setup_logger(__name__)


class ResultScraper:
    """Handles web scraping operations for VTU result portal"""
    
    def __init__(self, driver_path, website_url):
        """
        Initialize the scraper with driver path and website URL
        
        Args:
            driver_path: Path to ChromeDriver executable
            website_url: URL of the VTU result portal
        """
        self.driver_path = driver_path
        self.website_url = website_url
        self.driver = None
        self.actions = None
        self.main_window = None
        self.wait = None
        
        # Page elements
        self.usn_box = None
        self.captcha_box = None
        self.submit_btn = None
    
    def setup_driver(self):
        """Initialize and configure the Chrome WebDriver"""
        service = Service(self.driver_path)
        opts = Options()
        opts.add_experimental_option("detach", True)
        
        self.driver = webdriver.Chrome(service=service, options=opts)
        self.actions = ActionChains(self.driver)
        self.wait = WebDriverWait(self.driver, WAIT_TIMEOUT)
        self.driver.maximize_window()
        self.driver.get(self.website_url)
        
        self.main_window = self.driver.window_handles[0]
        logger.info(f"Driver setup complete. Opened {self.website_url}")
        # Initial wait for page load - can be kept or replaced by explicit wait for element
        time.sleep(WAIT_AFTER_STARTUP)
    
    def locate_page_elements(self):
        """
        Locate the main page elements (USN input, captcha input, submit button)
        
        Returns:
            True if all elements found, False otherwise
        """
        try:
            self.usn_box = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_USN_INPUT)))
            self.captcha_box = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_CAPTCHA_INPUT)))
            self.submit_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_SUBMIT_BUTTON)))
            logger.info("Page elements located successfully.")
            return True
        except (NoSuchElementException, TimeoutException) as e:
            logger.error(f"Page elements not found: {e}")
            return False
    
    def enter_usn_and_captcha(self, usn, captcha):
        """
        Enter USN and captcha values into the form
        
        Args:
            usn: Student USN to enter
            captcha: Captcha value to enter
        """

        self.usn_box.clear()
        self.usn_box.send_keys(usn)
        # Small sleep might still be needed for stability after typing, but reduced
        time.sleep(WAIT_AFTER_INPUT) 
        
        self.captcha_box.clear()
        self.captcha_box.send_keys(captcha)
        time.sleep(WAIT_AFTER_INPUT)
        logger.debug(f"Entered USN: {usn} and Captcha")
    
    def submit_and_switch_to_result(self):
        """
        Submit the form and switch to the result window
        
        Returns:
            True if result window opened successfully, False otherwise
        """
        # Open result in new tab using CTRL+Click
        self.actions.key_down(Keys.LEFT_CONTROL).perform()
        self.submit_btn.click()
        self.actions.key_up(Keys.LEFT_CONTROL).perform()
        
        # Wait for new window
        try:
            self.wait.until(EC.number_of_windows_to_be(2))
            result_window = self.driver.window_handles[1]
            self.driver.switch_to.window(result_window)
            logger.info("Switched to result window.")
            return True
        except TimeoutException:
            logger.warning("Result window did not open within timeout.")
            self.usn_box.clear()
            return False
    
    def scrape_student_info(self):
        """
        Scrape student USN and name from the result page
        
        Returns:
            Tuple of (usn, name) or (None, None) on error
        """
        try:
            usn_element = self.wait.until(EC.presence_of_element_located((By.XPATH, XPATH_STUDENT_USN)))
            name_element = self.driver.find_element(By.XPATH, XPATH_STUDENT_NAME)
            
            usn = usn_element.text
            name = name_element.text
            logger.info(f"Scraped info for USN: {usn}, Name: {name}")
            return usn, name
        except (NoSuchElementException, TimeoutException) as e:
            logger.error(f"Error extracting student info: {e}")
            return None, None
    
    def scrape_subjects(self):
        """
        Scrape all subject details from the result page
        
        Returns:
            List of dictionaries containing subject data (name, ia, see, total, res)
        """
        subjects = []
        
        for k in range(1, MAX_SUBJECTS + 1):
            try:
                # Build XPath for current subject row
                base_xpath = XPATH_SUBJECT_BASE.format(k + 1)
                
                # Extract subject information
                sub_name = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_NAME_INDEX}]"
                ).text
                ia = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_IA_INDEX}]"
                ).text
                see = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_SEE_INDEX}]"
                ).text
                total = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_TOTAL_INDEX}]"
                ).text
                res = self.driver.find_element(
                    By.XPATH, f"{base_xpath}/div[{SUBJECT_RESULT_INDEX}]"
                ).text
                
                subjects.append({
                    "name": sub_name,
                    "ia": ia,
                    "see": see,
                    "total": total,
                    "res": res
                })
            except NoSuchElementException:
                # No more subjects found
                break
        
        logger.info(f"Scraped {len(subjects)} subjects.")
        return subjects
    
    def close_result_and_return_to_main(self):
        """Close the result window and switch back to main window"""
        try:
            self.driver.close()
            self.driver.switch_to.window(self.main_window)
            self.usn_box.clear()
            # Wait for main window to be active/interactable again
            self.wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_USN_INPUT)))
            logger.debug("Closed result window and returned to main.")
        except Exception as e:
             logger.error(f"Error returning to main window: {e}")
    
    def cleanup(self):
        """Clean up driver resources"""

        if self.driver:
            try:
                self.driver.quit()
                logger.info("Driver exited successfully.")
            except Exception as e:
                logger.error(f"Error closing driver: {e}")
