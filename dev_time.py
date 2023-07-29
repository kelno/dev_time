# needs VirtualDesktopAccessor.dll to run - https://github.com/Ciantic/VirtualDesktopAccessor/raw/master/x64/Release/VirtualDesktopAccessor.dll
# SELECT * FROM total_work_time_per_week;

import time
import pyautogui
import pygetwindow
import sqlite3
import ctypes
from pynput import keyboard

# Global variables
last_mouse_position = None
conn = None
cursor = None
last_keyboard_event = None
keyboard_listener = None

# Time of script inactivity after which we should consider the data invalid. Helps with sleep/hibernate
INACTIVITY_THRESHOLD = 100

def on_press(key):
    # Store the last recorded keyboard event
    global last_keyboard_event
    last_keyboard_event = key

def has_keyboard_activity():
    # Check if there has been a keyboard event since the last check
    global last_keyboard_event
    current_event = last_keyboard_event
    last_keyboard_event = None
    if current_event is not None:
        print("Keyboard activity")
        
    return current_event is not None
    
def is_in_dev_env(vda):
    # seems I can't check by name before Windows 11 with that library. https://github.com/Ciantic/VirtualDesktopAccessor
    num = vda.GetCurrentDesktopNumber()
    return num == 1
    
def get_mouse_position():
    # Get the current mouse position
    return pyautogui.position()

def has_mouse_moved(current_position):
    # Check if the current position is different from the last recorded position
    global last_mouse_position
    if last_mouse_position is None:
        last_mouse_position = current_position
        return False
    elif last_mouse_position != current_position:
        last_mouse_position = current_position
        print("Detected mouse movement")
        return True
    else:
        return False

def detect_mouse_input_activity():
    current_position = get_mouse_position()
    if has_mouse_moved(current_position):
        return True
        
    return False
    
def detect_input_activity():
    mouse_activity = detect_mouse_input_activity()
    keyboard_activity = has_keyboard_activity() # always call this one even if mouse is detected, to clear last event
    return mouse_activity or keyboard_activity
    
def create_database_table():
    global conn
    global cursor
    
    # Connect to the SQLite database or create a new one if it doesn't exist
    conn = sqlite3.connect("work_time_data.db")
    cursor = conn.cursor()

    # Create the table if it doesn't exist
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS work_time (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            start_time TIMESTAMP NOT NULL,
            end_time TIMESTAMP NOT NULL,
            duration REAL NOT NULL
        )
    ''')
    
    cursor.execute('''
        DROP VIEW IF EXISTS total_work_time_per_week;
    ''')
    
    cursor.execute('''
        CREATE VIEW IF NOT EXISTS total_work_time_per_week AS
        SELECT
            strftime('%Y-%W', datetime(start_time, 'unixepoch')) AS week,
            ROUND(SUM(duration) / 3600.0, 1) AS total_duration_hours
        FROM
            work_time
        GROUP BY
            week;
    ''')

    # Commit the changes and close the connection
    conn.commit()

def insert_work_data(start_time, end_time, duration):
    global conn
    global cursor
    
    # Insert work time data into the table
    cursor.execute('''
        INSERT INTO work_time (start_time, end_time, duration)
        VALUES (?, ?, ?)
    ''', (start_time, end_time, duration))

    # Commit the changes and close the connection
    conn.commit()

def pretty_time(timestamp):
    time_struct = time.localtime(timestamp)
    pretty_time = time.strftime("%Y-%m-%d %H:%M:%S", time_struct)
    return pretty_time
    
def save_work(start_time):
    if start_time is None:
        return

    end_time = time.time()
    print(f"{pretty_time(end_time)}: Stopping work") 
    duration = end_time - start_time
    insert_work_data(start_time, end_time, duration)
    
def print_data():
    global conn
    global cursor
    
    cursor.execute('''
        SELECT * FROM total_work_time_per_week;
    ''')
    
    result = cursor.fetchall()
    
    # Print the result of the SELECT command
    print("\nTotal Work Time Per Week:")
    print("-------------------------")
    for row in result:
        print(f"{row[0]} | {row[1]} hours")

def main():
    global conn
    
    create_database_table()

    in_dev_mode = False
    start_time = None
    last_update_time = time.time()
    
    vda = ctypes.WinDLL("./VirtualDesktopAccessor.dll")

    keyboard_listener = keyboard.Listener(on_press=on_press)
    keyboard_listener.start()
    
    try:
        while True:
            current_time = time.time()
            # reset tracking if the script was inactive for a while
            time_difference = current_time - last_update_time
            if time_difference >= INACTIVITY_THRESHOLD:
                start_time = current_time
            
            is_input_active = detect_input_activity()

            if is_in_dev_env(vda) and is_input_active:
                if not in_dev_mode:
                    in_dev_mode = True
                    start_time = time.time()
                    print(f"{pretty_time(start_time)}: Starting work")
            else:
                if in_dev_mode: 
                    in_dev_mode = False
                    save_work(start_time)

            last_update_time = current_time
            time.sleep(60)
            
    except KeyboardInterrupt:
        print("\nKeyboardInterrupt: Saving data and exiting...")
        save_work(start_time)
        print_data()
        conn.close()

if __name__ == "__main__":
    main()
    input("Press any key to exit...")
