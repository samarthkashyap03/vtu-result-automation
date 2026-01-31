
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


def process_results(gui, inputs):
    """
    Main automation workflow - processes student results
    
    Args:
        gui: AutomationGUI instance for user interaction
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
        gui.show_error("Start/End row must be numbers.")
        return
    
    # Load input Excel file containing USNs
    in_book, in_sheet = load_input_workbook(usn_file)
    if not in_book:
        gui.show_error("Failed to load input workbook.")
        return
    
    # Create output Excel workbook for results
    out_book, out_sheet, orange_style = create_output_workbook()
    
    # Write static headers
    write_headers(out_sheet)
    
    time.sleep(WAIT_AFTER_STARTUP)
    
    # Initialize web scraper
    scraper = ResultScraper(driver_path, website)
    try:
        scraper.setup_driver()
    except Exception as e:
        gui.show_error(f"Failed to setup driver: {e}")
        return
    
    # Locate page elements
    if not scraper.locate_page_elements():
        gui.show_error("Page elements not found. Website layout may have changed.")
        return
    
    time.sleep(WAIT_BEFORE_CAPTCHA)
    
    # Get captcha from user
    captcha = gui.get_captcha_input()
    if not captcha:
        gui.show_warning("Captcha input cancelled. Stopping.")
        scraper.cleanup()
        return
    
    # Process each USN from the input file
    success_count = 0
    error_count = 0
    
    for row in range(start_row, end_row + 1):
        usn = read_usn(in_sheet, row)
        
        if not usn:
            continue
        
        # Enter USN and captcha, submit the form
        scraper.enter_usn_and_captcha(usn, captcha)
        
        # Open result page in new window
        if not scraper.submit_and_switch_to_result():
            error_count += 1
            continue
        
        try:
            # Extract student information
            page_usn, name = scraper.scrape_student_info()
            if not page_usn or not name:
                error_count += 1
                scraper.close_result_and_return_to_main()
                continue
            
            # Extract all subject details
            subjects = scraper.scrape_subjects()
            
            # Write student info to Excel
            write_student_info(out_sheet, row, page_usn, name)
            
            # Write subject data to Excel
            write_subject_data(out_sheet, row, subjects, orange_style)
            success_count += 1
            
        except Exception as e:
            print(f"Error processing {usn}: {e}")
            error_count += 1
        
        # Close result window and return to main page
        scraper.close_result_and_return_to_main()
    
    # Save the output Excel file
    save_workbook(out_book, save_path)
    
    gui.show_info(f"Processing complete.\nProcessed: {success_count}\nErrors: {error_count}")
    scraper.cleanup()


def main():
    """Application entry point - creates GUI and starts the workflow"""
    
    # Create GUI with callback to process_results
    # We pass the gui instance itself to process_results so it can use dialog methods
    # Note: we can't pass 'gui' directly in the lambda because it's not defined yet.
    # We define gui first.
    
    gui = None # placeholder
    
    def on_submit():
        if gui:
            process_results(gui, gui.get_inputs())

    gui = AutomationGUI(on_submit_callback=on_submit)
    
    # Start the GUI
    gui.run()


if __name__ == "__main__":
    main()
