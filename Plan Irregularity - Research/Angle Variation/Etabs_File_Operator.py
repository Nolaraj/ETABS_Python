import pyautogui
from time import sleep
import os
from tkinter import Tk, filedialog

#Install Pyautogui, opencv_python, pillow and tkinter packages

timer_a=2
print(f'Swiftly open the browser window showing required icons')
sleep(timer_a)


#Extract Folder path from the dialog box
def folder_path(title="Enter the file path"):
    root = Tk()             # pointing root to Tk() to use it as Tk() in program.
    root.withdraw()         # Hides small tkinter window.
    root.title()
    root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.
    # root.attributes()
    open_file = filedialog.askdirectory(title = title) #,filetypes = (("Excel files", ".xlsx .xls"),))   # Returns opened path as str
    return open_file
icons_file = folder_path()
os.chdir(icons_file)


def image_checkpoint_return_location(image, confidence = 0.5):
    try:
        location = pyautogui.locateOnScreen(image, confidence=confidence) #For Confidence to be used in Code pip install opencv-python is needed
        if location is not None:
            return (location)
        else:
            print(f'{image} position is None during runtime')
            return (location)
    except:
        print(f'{image} is not available in the screen during runtime, So Scrolling Down')
        return("Scroll Down")




def Wait_Until_Image_Appears_and_Click(image, confidence = 0.5):
    while image_checkpoint_return_location(image) is None:
        sleep(1)
    pyautogui.click(image_checkpoint_return_location(image, confidence))

def Sleep_Until_Image_Disappears(image):
    while image_checkpoint_return_location(image) != None:
        sleep(1)



while image_checkpoint_return_location('rows_per_page.png', 0.9) == None:
    print(image_checkpoint_return_location('edit_draft.png', 0.96))


    if image_checkpoint_return_location('edit_draft.png') == None:
        # for i in range(1, 20):
        pyautogui.scroll(-200)
        print("Scrolling")

    else:
        # Wait_Until_Image_Appears_and_Click('edit_draft.png', 0.96)
        sleep(1)

        pyautogui.click(image_checkpoint_return_location('edit_draft.png'))

        Wait_Until_Image_Appears_and_Click('visibility_button.png')
        Wait_Until_Image_Appears_and_Click('private_radio.png')
        pyautogui.click(image_checkpoint_return_location('save_button.png', 0.9))
        Sleep_Until_Image_Disappears("info_close.png")
        sleep(1)
