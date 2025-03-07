import os
import time
import pyautogui
import win32gui
import win32con
import pytesseract
import re
from PIL import Image, ImageGrab
import cv2
import numpy as np
from datetime import datetime

# Configure pytesseract path - update this to your installation path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Time to wait between checks (in seconds)
WAIT_TIME = 1
TIMEOUT = 30  # Timeout after this many seconds
VIBER_WINDOW_TITLE = "Rakuten Viber"  # Title of the Viber window

# Output directories
OUTPUT_DIR = r"C:\Users\Admin\Documents\Viber_Attachments"
SCREENSHOT_DIR = os.path.join(OUTPUT_DIR, "Screenshots")

def ensure_directories_exist():
    """Create necessary directories if they don't exist"""
    try:
        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            print(f"Created output directory: {OUTPUT_DIR}")
            
        if not os.path.exists(SCREENSHOT_DIR):
            os.makedirs(SCREENSHOT_DIR)
            print(f"Created screenshot directory: {SCREENSHOT_DIR}")
    except Exception as e:
        print(f"Error creating directories: {e}")

def bring_viber_to_foreground():
    """Find the Viber window and bring it to the foreground in full screen"""
    try:
        # Find the Viber window by title
        hwnd = win32gui.FindWindow(None, VIBER_WINDOW_TITLE)
        
        if hwnd:
            # Maximize the window
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            # Bring it to the foreground
            win32gui.SetForegroundWindow(hwnd)
            # Give UI time to update
            time.sleep(0.5)
            return True
        else:
            print("Viber window not found")
            return False
    except Exception as e:
        print(f"Error bringing Viber to foreground: {e}")
        return False

def capture_viber_screenshot():
    """Capture a screenshot of the Viber window"""
    try:
        # Take a screenshot of the entire screen
        screenshot = ImageGrab.grab()
        return screenshot
    except Exception as e:
        print(f"Error capturing screenshot: {e}")
        return None

def extract_specific_regions(screenshot):
    """Extract and process specific regions of the screenshot for better OCR accuracy"""
    try:
        # Convert PIL image to numpy array for OpenCV
        screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        
        # Get image dimensions
        height, width = screenshot_cv.shape[:2]
        
        # Define regions of interest (ROIs) based on the screenshots provided
        # These values are percentages of the total screen size
        rois = {
            'header': {
                'top': 0.08,  # 8% from top
                'left': 0.4,  # 40% from left
                'bottom': 0.15,  # 15% from top
                'right': 0.85  # 85% from left
            },
            'transaction_info': {
                'top': 0.3,  # 30% from top
                'left': 0.3,  # 30% from left
                'bottom': 0.7,  # 70% from top
                'right': 0.85  # 85% from left
            },
            'bank_details': {
                'top': 0.6,  # 60% from top
                'left': 0.4,  # 40% from left
                'bottom': 0.9,  # 90% from top
                'right': 0.9  # 90% from left
            }
        }
        
        # Extract and OCR each region
        region_texts = {}
        for region_name, coords in rois.items():
            # Calculate pixel coordinates
            top = int(height * coords['top'])
            left = int(width * coords['left'])
            bottom = int(height * coords['bottom'])
            right = int(width * coords['right'])
            
            # Extract region
            roi = screenshot_cv[top:bottom, left:right]
            
            # Save region for debugging if needed
            region_path = os.path.join(SCREENSHOT_DIR, f"debug_{region_name}.png")
            cv2.imwrite(region_path, roi)
            
            # Process the region for better OCR
            # Convert to grayscale
            roi_gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
            
            # Apply slight blur to reduce noise
            roi_blur = cv2.GaussianBlur(roi_gray, (3, 3), 0)
            
            # Apply adaptive thresholding to enhance text
            roi_thresh = cv2.adaptiveThreshold(roi_blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                             cv2.THRESH_BINARY, 11, 2)
            
            # Perform OCR on the processed region
            region_text = pytesseract.image_to_string(roi_thresh)
            region_texts[region_name] = region_text
            print(f"OCR for {region_name}: {region_text}")
            
        return region_texts
    except Exception as e:
        print(f"Error extracting regions: {e}")
        return {}

def extract_info_from_screenshot(screenshot):
    """Extract contact name, number, and date/time from the screenshot"""
    try:
        # First try region-specific extraction for better accuracy
        region_texts = extract_specific_regions(screenshot)
        
        # Combine all region texts for comprehensive search
        all_text = "\n".join(region_texts.values()) if region_texts else ""
        
        # If region extraction failed, fall back to full image OCR
        if not all_text:
            # Convert PIL image to cv2 format for processing
            screenshot_cv = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            # Convert to grayscale for better OCR results
            gray = cv2.cvtColor(screenshot_cv, cv2.COLOR_BGR2GRAY)
            
            # Use pytesseract to extract all text from the image
            all_text = pytesseract.image_to_string(gray)
        
        print(f"OCR extracted text: {all_text}")  # Debug: print all extracted text
        
        # Extract contact info using regex patterns based on the screenshots
        # Pattern for the contact name (based on the screenshots showing "SoneeHardware JanavareeMagu")
        contact_name_match = re.search(r"Sonee\w+\s+\w+", all_text)
        contact_name = contact_name_match.group(0) if contact_name_match else "Unknown"
        
        # Extract bank info or account numbers visible in the screenshot (MVR 7701 70007B 001)
        account_match = re.search(r"MVR\s+\d+\s+\d+\w*\s+\d+|77017000\d+", all_text)
        account_number = account_match.group(0).replace(' ', '') if account_match else ""
        
        # Look for transaction amount (like 1,065.80)
        amount_match = re.search(r"\d{1,3}(?:,\d{3})*\.\d{2}", all_text)
        amount = amount_match.group(0).replace(',', '') if amount_match else ""
        
        # Pattern for date (Monday, February 24, 2025)
        date_match = re.search(r"(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}", all_text)
        date_str = date_match.group(0) if date_match else ""
        
        # Extract time (Last seen today at 4:24 PM)
        time_match = re.search(r"Last\s+seen\s+today\s+at\s+(\d{1,2}:\d{2}\s+(?:AM|PM))", all_text)
        time_str = time_match.group(1) if time_match else ""
        
        # Look for transaction date (24/02/2025)
        transaction_date_match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", all_text)
        transaction_date = transaction_date_match.group(0) if transaction_date_match else ""
        
        # Create a properly formatted date/time string
        if date_str and time_str:
            date_time = f"{date_str} {time_str}"
        elif date_str:
            date_time = date_str
        elif time_str:
            date_time = f"Today {time_str}"
        elif transaction_date:
            date_time = transaction_date
        else:
            date_time = time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime())
            
        # Extract reference number if present (e.g., BLAZ29Y417504782)
        ref_match = re.search(r"Reference\s+([A-Z0-9]+)", all_text, re.IGNORECASE)
        reference = ref_match.group(1) if ref_match else ""
        
        # Also look for "successful" transaction confirmation
        success_match = re.search(r"(SUCCESS|transaction\s+is\s+successful)", all_text, re.IGNORECASE)
        success_status = "SUCCESS" if success_match else ""
        
        # Clean up the extracted information (remove invalid characters for filenames)
        contact_name = re.sub(r'[<>:"/\\|?*]', '', contact_name).strip()
        if not contact_name or contact_name == "Unknown":
            # Try an alternative pattern like looking for "To SONEE" in the transaction details
            alt_contact_match = re.search(r"To\s+(SONEE\s+\w+)", all_text)
            if alt_contact_match:
                contact_name = alt_contact_match.group(1).strip()
                
        # Combine account number, amount, reference, and status for a more detailed filename
        details = []
        if amount:
            details.append(f"MVR{amount}")
        if account_number:
            details.append(account_number)
        if reference:
            details.append(reference)
        if success_status:
            details.append(success_status)
            
        details_str = "_".join(details) if details else "NoDetails"
        
        # Format date_time to be filename friendly
        date_time = re.sub(r'[<>:"/\\|?*,]', '_', date_time).strip()
        
        print(f"Extracted info - Name: {contact_name}, Details: {details_str}, Date: {date_time}")
        return contact_name, details_str, date_time
    except Exception as e:
        print(f"Error extracting info from screenshot: {e}")
        return "Unknown", "NoDetails", time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime())

def wait_for_final_file(tmp_path):
    """ Wait until the temporary file is completed and renamed to its final form """
    try:
        # Remove '_tmp' suffix to form the final expected filename
        base_name = tmp_path.rstrip('_tmp')

        # Ensure the file exists and keep checking until it's stable
        start_time = time.time()

        while True:
            if os.path.exists(base_name):
                # Check if the file has stopped changing (by checking its modification time)
                current_mod_time = os.path.getmtime(base_name)
                time.sleep(WAIT_TIME)  # Wait a little before checking again
                new_mod_time = os.path.getmtime(base_name)
                
                if new_mod_time == current_mod_time:
                    print(f"Final file found and stable: {base_name}")
                    # Bring Viber to foreground before renaming
                    if bring_viber_to_foreground():
                        screenshot = capture_viber_screenshot()
                        if screenshot:
                            # Process the screenshot to get info for renaming
                            rename_file(base_name, screenshot)
                        else:
                            # Fallback to basic rename if screenshot fails
                            rename_file(base_name)
                    else:
                        # Fallback to basic rename if bringing to foreground fails
                        rename_file(base_name)
                    break
                else:
                    print(f"File still being written: {base_name}")
            else:
                # Check if the timeout has passed
                if time.time() - start_time > TIMEOUT:
                    print(f"Timeout reached. Final file {base_name} not found.")
                    break

            # Wait before the next check
            time.sleep(WAIT_TIME)
    except Exception as e:
        print(f"Error waiting for file completion: {e}")

def rename_file(original_path, screenshot=None):
    """ Renames the file to the desired format using information from screenshot if available """
    filename = os.path.basename(original_path)
    print(f"Attempting to rename file: {filename}")  # Debugging line

    # Skip files that have already been renamed by checking for timestamp pattern in the filename
    if "_tmp" in filename or (filename.count("_") >= 2 and re.search(r'\d{4}-\d{2}-\d{2}', filename)):
        print(f"Skipping file (already processed or temporary): {filename}")
        return

    # Get file extension
    _, file_extension = os.path.splitext(original_path)
    
    # Current timestamp for unique filenames
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if screenshot:
        # Extract contact name, number, and date from the screenshot
        contact_name, details_str, date_time = extract_info_from_screenshot(screenshot)
        
        # Format date_time to be filename friendly
        formatted_date = re.sub(r'[/,\s:]', '_', date_time)
        
        # Create new filename with extracted information
        new_name = f"{contact_name}_{details_str}_{formatted_date}{file_extension}"
        
        # Save a copy of the screenshot with a similar name for reference
        screenshot_name = f"{contact_name}_{details_str}_{formatted_date}_screenshot_{timestamp}.png"
        screenshot_path = os.path.join(SCREENSHOT_DIR, screenshot_name)
        try:
            # Ensure the screenshot directory exists
            ensure_directories_exist()
            # Save the screenshot
            screenshot.save(screenshot_path)
            print(f"Screenshot saved to: {screenshot_path}")
        except Exception as e:
            print(f"Error saving screenshot: {e}")
    else:
        # Fallback to simple renaming logic if no screenshot is available
        name_parts = filename.split('_')
        
        if len(name_parts) > 1:
            new_name = f"{name_parts[0]}_{name_parts[1]}_{timestamp}{file_extension}"
        else:
            new_name = f"Unknown_Unknown_{timestamp}{file_extension}"

    # Define new file path
    new_path = os.path.join(OUTPUT_DIR, new_name)

    # Check if the file already has the desired name
    if original_path == new_path:
        print(f"File already has the correct name: {original_path}")
        return

    # Ensure the output directory exists
    ensure_directories_exist()

    # Rename the file
    try:
        print(f"Renaming file: {original_path} -> {new_path}")
        os.rename(original_path, new_path)
    except PermissionError as e:
        print(f"Error renaming file: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

def watch_viber_folder(folder_path):
    """ Monitors the folder for new files and processes them """
    # Create a dictionary to track files and their status
    file_tracking = {}
    
    # Get the initial list of files in the directory
    try:
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            file_tracking[file_path] = {
                'processed': True,  # Mark as processed since they existed before script started
                'time_detected': time.time(),
                'size': os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        print(f"Initialized tracking for {len(file_tracking)} existing files")
    except Exception as e:
        print(f"Error initializing file tracking: {e}")
        file_tracking = {}

    print(f"Watching folder: {folder_path}")
    print("Press Ctrl+C to stop the script")
    
    try:
        while True:
            try:
                # Check the folder for new or modified files
                current_files = set()
                for filename in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, filename)
                    current_files.add(file_path)
                    
                    # Skip temporary zero-byte files
                    if os.path.getsize(file_path) == 0:
                        continue
                    
                    # Check if this is a new file or if the file has been modified
                    if file_path not in file_tracking:
                        # New file detected
                        print(f"New file detected: {file_path}")
                        file_tracking[file_path] = {
                            'processed': False,
                            'time_detected': time.time(),
                            'size': os.path.getsize(file_path)
                        }
                    else:
                        # Check if the file has changed size and was previously processed
                        current_size = os.path.getsize(file_path)
                        if current_size != file_tracking[file_path]['size'] and file_tracking[file_path]['processed']:
                            print(f"File modified: {file_path}")
                            file_tracking[file_path] = {
                                'processed': False, 
                                'time_detected': time.time(),
                                'size': current_size
                            }
                
                # Process any unprocessed files
                for file_path, status in file_tracking.items():
                    if not status['processed'] and os.path.exists(file_path):
                        try:
                            filename = os.path.basename(file_path)
                            
                            # Check if file has stabilized (size hasn't changed for a few seconds)
                            current_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                            if current_size != status['size']:
                                # File is still being written, update size and skip for now
                                file_tracking[file_path]['size'] = current_size
                                continue
                                
                            # Process different file types
                            if "_tmp" in filename:
                                print(f"Processing temporary file: {file_path}")
                                wait_for_final_file(file_path)
                            else:
                                print(f"Processing file: {file_path}")
                                # For non-temporary files, bring Viber to foreground and take screenshot
                                if bring_viber_to_foreground():
                                    screenshot = capture_viber_screenshot()
                                    if screenshot:
                                        rename_file(file_path, screenshot)
                                    else:
                                        rename_file(file_path)
                                else:
                                    rename_file(file_path)
                            
                            # Mark as processed
                            file_tracking[file_path]['processed'] = True
                            
                        except Exception as e:
                            print(f"Error processing file {file_path}: {e}")
                            # If we encountered an error, wait longer before retrying
                            time.sleep(2)
                
                # Remove tracking for files that no longer exist
                missing_files = [f for f in file_tracking if f not in current_files]
                for file_path in missing_files:
                    print(f"File no longer exists, removing from tracking: {file_path}")
                    del file_tracking[file_path]
                
                # Sleep a bit before checking again
                time.sleep(2)
                
            except Exception as e:
                print(f"Error in main watch loop: {e}")
                time.sleep(5)  # Sleep longer after errors
    
    except KeyboardInterrupt:
        print("\nScript stopped by user")
    except Exception as e:
        print(f"Fatal error: {e}")
    finally:
        print("Exiting file watch")

# Folder to watch for new Viber downloads
folder_to_watch = r"C:\Users\Admin\Documents\ViberDownloads"

if __name__ == "__main__":
    print("Starting Viber file renaming script with screenshot capability...")
    
    # Create necessary directories
    ensure_directories_exist()
    
    # Check if required libraries are available
    try:
        import pyautogui
        import win32gui
        import pytesseract
        from PIL import ImageGrab
        import cv2
        print("All required libraries loaded successfully.")
        
        # Notify user about Tesseract OCR configuration
        print("Note: Make sure Tesseract OCR is installed and the path is correctly set in the script.")
        print(f"Current Tesseract path: {pytesseract.pytesseract.tesseract_cmd}")
        
    except ImportError as e:
        print(f"Error: Missing required library - {e}")
        print("Please install missing libraries using pip:")
        print("pip install pyautogui pywin32 pytesseract pillow opencv-python")
        exit(1)
        
    watch_viber_folder(folder_to_watch)
