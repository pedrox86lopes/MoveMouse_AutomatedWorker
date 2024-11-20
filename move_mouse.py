import pyautogui
import random
import time
import openpyxl
import os
import subprocess

def create_and_open_excel(filename="activity_data.xlsx"):
    """
    Creates an Excel file and opens it using the default application.
    """
    # Create the Excel file
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "ActivityData"
    sheet.append(["Index", "Value"])  # Add headers
    workbook.save(filename)
    print(f"Excel file '{filename}' created.")

    # Open the Excel file with the default application
    if os.name == "nt":  # Windows
        os.startfile(filename)
    elif os.name == "posix":  # macOS and Linux
        subprocess.call(["open", filename])

    return filename

def simulate_typing_in_excel():
    """
    Simulates typing data into the Excel file.
    """
    # Randomize starting cell for typing
    pyautogui.click(x=random.randint(100, 300), y=random.randint(200, 400))  # Click somewhere in the Excel window
    for _ in range(random.randint(5, 10)):  # Type a random number of rows
        random_number = random.randint(1, 1000)
        pyautogui.typewrite(str(random_number))
        pyautogui.press("enter")
        time.sleep(random.uniform(0.5, 2.0))  # Delay between rows

def move_mouse_randomly():
    """
    Moves the mouse to random positions on the screen.
    """
    screen_width, screen_height = pyautogui.size()
    x = random.randint(0, screen_width - 1)
    y = random.randint(0, screen_height - 1)
    pyautogui.moveTo(x, y, duration=random.uniform(0.5, 2.0))

def perform_random_actions(filename="activity_data.xlsx"):
    """
    Alternates between typing in the open Excel file and moving the mouse randomly.
    """
    print("Performing random actions... Press Ctrl+C to stop.")
    while True:
        action = random.choice(["type_in_excel", "move_mouse"])
        if action == "type_in_excel":
            print("Typing in Excel...")
            simulate_typing_in_excel()
        elif action == "move_mouse":
            print("Moving mouse randomly...")
            move_mouse_randomly()
        time.sleep(random.uniform(5, 10))  # Random delay between actions

if __name__ == "__main__":
    try:
        # Step 1: Create and open the Excel file
        excel_file = create_and_open_excel()

        # Step 2: Perform random actions
        perform_random_actions(excel_file)
    except KeyboardInterrupt:
        print("\nSimulation stopped by user.")
