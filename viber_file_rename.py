import os
import time

# Time to wait between checks (in seconds)
WAIT_TIME = 1
TIMEOUT = 30  # Timeout after this many seconds

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
                    rename_file(base_name)  # Renaming the file once it is stable
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

def rename_file(original_path):
    """ Renames the file to the desired format """
    filename = os.path.basename(original_path)
    print(f"Attempting to rename file: {filename}")  # Debugging line

    # Skip files that have already been renamed by checking for timestamp pattern in the filename
    if "_tmp" in filename or filename.count("_") < 2:
        print(f"Skipping file (already processed or temporary): {filename}")
        return

    # Try to extract the desired name
    name_parts = filename.split('_')

    if len(name_parts) > 1:
        timestamp = time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime())
        new_name = f"{name_parts[0]}_{name_parts[1]}_{timestamp}.jpg"
    else:
        new_name = f"Unknown_Unknown_{time.strftime('%Y-%m-%d_%H-%M-%S', time.localtime())}.jpg"

    # Define new file path
    new_path = os.path.join("C:\\Users\\Admin\\Documents\\Viber_Attachments", new_name)

    # Check if the file already has the desired name
    if original_path == new_path:
        print(f"File already has the correct name: {original_path}")
        return

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
    # Get the initial list of files in the directory
    existing_files = set(os.path.abspath(os.path.join(folder_path, filename)) for filename in os.listdir(folder_path))

    while True:
        try:
            # Check the folder for new files
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                
                if os.path.abspath(file_path) in existing_files:
                    continue  # Skip files already processed (existing before the script)

                print(f"New file detected: {file_path}")

                if "_tmp" in filename:
                    print(f"Temporary file detected: {file_path}")
                    wait_for_final_file(file_path)  # Wait for the final file
                else:
                    rename_file(file_path)  # Rename the file
                    existing_files.add(os.path.abspath(file_path))  # Add to processed files list

            # Sleep a bit before checking again
            time.sleep(5)
        except Exception as e:
            print(f"Error while watching folder: {e}")
            time.sleep(5)

# Folder to watch for new Viber downloads
folder_to_watch = r"C:\Users\Admin\Documents\ViberDownloads"

if __name__ == "__main__":
    print("Starting file renaming script...")
    watch_viber_folder(folder_to_watch)
