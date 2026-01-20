
# VTU Result Automation - Main Script

import time
from config import WAIT_AFTER_STARTUP, WAIT_BEFORE_CAPTCHA
from gui import AutomationGUI
from scraper import ResultScraper
from excel_io import (
    load_input_workbook,
    read_usn,
    create_output_workbook,
    write_headers,
    write_student_info,
    write_subject_data,
    save_workbook
)


def process_results(inputs):
    """
    Main automation workflow - processes student results
    
    Args:
        inputs: Dictionary containing user inputs from GUI
    """
    driver_path = inputs["driver_path"]
    usn_file = inputs["usn_file"]
    website = inputs["website"]
    save_path = inputs["save_path"]
    
    # Validate row inputs
    try:
        start_row = int(inputs["start_row"])
        end_row = int(inputs["end_row"])
    except ValueError:
        print("Start/End row must be numbers.")
        return
    
    # Load input Excel file containing USNs
    in_book, in_sheet = load_input_workbook(usn_file)
    if not in_book:
        return
    
    # Create output Excel workbook for results
    out_book, out_sheet, orange_style = create_output_workbook()
    
    # Write static headers
    write_headers(out_sheet)
    
    time.sleep(WAIT_AFTER_STARTUP)
    
    # Initialize web scraper
    scraper = ResultScraper(driver_path, website)
    scraper.setup_driver()
    
    # Locate page elements
    if not scraper.locate_page_elements():
        return
    
    time.sleep(WAIT_BEFORE_CAPTCHA)
    
    # Get captcha from user
    captcha = input("Enter captcha shown in the browser: ")
    
    print(f"Processing rows {start_row} to {end_row}...")
    
    # Process each USN from the input file
    for row in range(start_row, end_row + 1):
        usn = read_usn(in_sheet, row)
        
        if not usn:
            print(f"Skipping row {row}: empty USN")
            continue
        
        # Enter USN and captcha, submit the form
        scraper.enter_usn_and_captcha(usn, captcha)
        
        # Open result page in new window
        if not scraper.submit_and_switch_to_result():
            continue
        
        try:
            # Extract student information
            page_usn, name = scraper.scrape_student_info()
            if not page_usn or not name:
                continue
            
            # Extract all subject details
            subjects = scraper.scrape_subjects()
            
            # Write student info to Excel
            write_student_info(out_sheet, row, page_usn, name)
            
            # Write subject data to Excel
            write_subject_data(out_sheet, row, subjects, orange_style)
            
        except Exception as e:
            print(f"Error processing {usn}: {e}")
        
        # Close result window and return to main page
        scraper.close_result_and_return_to_main()
    
    # Save the output Excel file
    save_workbook(out_book, save_path)


def main():
    """Application entry point - creates GUI and starts the workflow"""
    
    # Create GUI with callback to process_results
    gui = AutomationGUI(on_submit_callback=lambda: process_results(gui.get_inputs()))
    
    # Start the GUI
    gui.run()


if __name__ == "__main__":
    main()
