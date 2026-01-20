"""Configuration file for VTU Result Automation
contains XPath selectors, constants, and configuration values
"""

# XPath selectors for the main result page
XPATH_USN_INPUT = '//*[@id="raj"]/div[1]/div/input'
XPATH_CAPTCHA_INPUT = '//*[@id="raj"]/div[2]/div[1]/input'
XPATH_SUBMIT_BUTTON = '//*[@id="submit"]'

# XPath selectors for result page data extraction
XPATH_STUDENT_USN = '//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]'
XPATH_STUDENT_NAME = '//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]'
XPATH_SUBJECT_BASE = '//*[@id="dataPrint"]/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{}]'

# Excel configuration
EXCEL_HEADER_ROW = 0
EXCEL_SUBHEADER_ROW = 1
EXCEL_USN_COLUMN = 0
EXCEL_NAME_COLUMN = 2
EXCEL_SUBJECTS_START_COLUMN = 4

# Subject columns in result page
SUBJECT_NAME_INDEX = 1
SUBJECT_IA_INDEX = 3
SUBJECT_SEE_INDEX = 4
SUBJECT_TOTAL_INDEX = 5
SUBJECT_RESULT_INDEX = 6

# Maximum number of subjects to check
MAX_SUBJECTS = 14

# Wait times (in seconds)
WAIT_AFTER_STARTUP = 3
WAIT_BEFORE_CAPTCHA = 3
WAIT_AFTER_INPUT = 0.5
