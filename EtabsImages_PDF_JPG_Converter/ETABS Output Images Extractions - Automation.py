import pdf2image
import pyautogui, time
import os
from tkinter import Tk, filedialog
from PIL import Image
from time import sleep
import os
import pdf2image


normal_sleep = 2
window_transition = 1
bays_X = 5
bays_Y = 2
storey_Z = 5
elevation_grids = bays_X+1+bays_Y+1
plan_grids = storey_Z+1


# Extract path from the file dialog box
def path_finder(title):
    root = Tk() # pointing root to Tk() to use it as Tk() in program.
    root.withdraw() # Hides small tkinter window.
    root.title()
    root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.
    # root.attributes()
    open_file = filedialog.askdirectory(title = title) #,filetypes = (("Excel files", ".xlsx .xls"),))   # Returns opened path as str
    return open_file
def image_checkpoint_return_location(image):
    try:
        location = pyautogui.locateOnScreen(image, confidence = 0.5)
        if location is not None:
            print(location)
            return (location)
        else:
            print(f'{image} postition is None during runtime')
            #Active value not set as per required by code error
            if image == "longitudinal_reinforcing.png":
                print(f'Set {image} as active before running the code')
            return (eroor)
    except:

        print(f'{image} is not available in the screen during runtime')
def click_right_one_third(box):
    x = int(box[0]) + int(box[2])*0.83
    y = int(box[1]) + int(box[3]) / 2
    pyautogui.click(x,y)
def click_left_one_third(box):
    x = int(box[0]) + int(box[2])*0.167
    y = int(box[1]) + int(box[3]) / 2
    pyautogui.click(x,y)
def hold_on(icon):
    presence = pyautogui.locateOnScreen(icon)
    while  presence != "None":
        sleep(1)
        presence = pyautogui.locateOnScreen(icon)
        if presence == None:
            sleep(1)
            break

def printer(property,view, index):
    pyautogui.hotkey('ctrl', 'p');    sleep(6)
    pyautogui.click(image_checkpoint_return_location('print.png'));    sleep(1)
    hold_on("print_hold.png")
    pyautogui.click(image_checkpoint_return_location('file_name.png'));
    pyautogui.write(f'{property} {view} {index}'); pyautogui.press('tab', presses=3); pyautogui.press('enter'); sleep(10)
    pyautogui.hotkey('alt', 'f4'); sleep(2)

os.chdir(path_finder("Enter the folder of icons"))
pyautogui.doubleClick(image_checkpoint_return_location('3d_view.png'))

# print(pyautogui.locateOnScreen("click_to_cancel.png"))
# Concrete Frame Design_____________________________________________________________________________________________
# click_left_one_third(image_checkpoint_return_location('concrete_frame_design.png'))
# hold_on("click_to_cancel.png")

# with pyautogui.hold('shift'):
#     with pyautogui.hold('ctrl'):
#         pyautogui.press('f6')

# Concrete Design Parameters________________________________________________________________________________________
pyautogui.hotkey('shift', 'ctrl', 'f6')
sleep(window_transition)
# pyautogui.doubleClick( )

index_viewdata = {}


def design_details():
    sleep(2)
    box = image_checkpoint_return_location('design_results_entry.png')
    a = pyautogui.center(box)
    im = Image.open('other_outputs.png')
    width, height = im.size
    unit_height = height/10
    base_y = a[1] + height + 1.2*unit_height
    x = a[0]
    y = []
    [y.append(a[1] + unit_height*i) for i in range(1,11)]
    for i in range(0, 10):
        pyautogui.hotkey('shift', 'ctrl', 'f6')
        sleep(1)
        pyautogui.click(a[0], a[1])
        pyautogui.click(x, y[i])
        pyautogui.hotkey('tab', 'enter')

        if i in [0,1,2]: #For Plan and Elevation data extraction
            #Elevation Screenshot
            pyautogui.hotkey('ctrl', 'shift', 'f2'); sleep(1)
            pyautogui.press('tab', presses=6);  pyautogui.press('enter'); sleep(2)
            printer(i,"Elevation",  1)
            for j in range(0,elevation_grids-1):
                pyautogui.click(image_checkpoint_return_location('up_arrow.png')); sleep(2)
                printer(i, "Elevation", j+2)

            #Plan Screenshot
            pyautogui.hotkey('ctrl', 'shift', 'f1'); sleep(1)
            pyautogui.click(image_checkpoint_return_location('base.png'));
            pyautogui.click(image_checkpoint_return_location('ok.png'));

            sleep(2)
            printer(i, "Plan" ,1)
            for j in range(0,plan_grids-1):
                pyautogui.click(image_checkpoint_return_location('up_arrow.png')); sleep(2)
                printer(i,"Plan", j+1)

        else: #For 3D Data Extraction
            pyautogui.doubleClick(image_checkpoint_return_location('3d_view.png'))
            printer(i, "3D", 1)
            pyautogui.hotkey('ctrl', 'shift', 'f3'); sleep(1)
            pyautogui.click(image_checkpoint_return_location('fast_3d.png'))
            pyautogui.press('tab', presses = 8); pyautogui.press('enter', presses = 36)
            pyautogui.press('tab', presses=3); pyautogui.press('enter', presses=16)
            pyautogui.press('tab', presses=3); pyautogui.press('enter', presses=6)
            pyautogui.press('tab', presses=5);pyautogui.press('enter')
            printer(i, "3D", 2)

# pyautogui.click(image_checkpoint_return_location('down_arrow.png'))
    pyautogui.hotkey('shift', 'ctrl', 'f6')
    pyautogui.click(a[0], a[1])
    pyautogui.click(x, y[0])
    pyautogui.hotkey('tab', 'enter')
#design_details()


#Importing pdf and saving to Image Jpeg
# pdf_directory = path_finder("Extracted Pdf folder")
pdf_directory = 'F:/Nolaraj/Practice Stuff/Images'
for items in os.listdir(pdf_directory):
    pages =  pdf2image.convert_from_path(items)
    for i in range(len(pages)):
        pages[i].save(f'{items[:-4]}.jpg', 'JPEG')















#parameters = ['bcc_ratio.png','general_reinforcements.png', 'column_pmm.png', 'cbc_ratio.png',  'identify_all.png', 'rebar_percentage.png','identify_pm.png', 'shear_reinforcing.png',  'longitudinal_reinforcing.png','identify_shear.png', ]
# for parameter in parameters:
#    pyautogui.hotkey('shift', 'ctrl', 'f6')
#    sleep(4)
#    pyautogui.click(a[0], a[1])
#    sleep(1)
#    pyautogui.moveRel(0,-100)
#    pyautogui.click(image_checkpoint_return_location(parameter))
#    sleep(1)
#    pyautogui.hotkey('tab', 'enter')
#    sleep(2)
