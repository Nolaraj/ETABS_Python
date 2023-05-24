import pyautogui, time
import os
from PIL import Image
from time import sleep
import os
import main
from win32com import client
os.chdir(main.etabs_icon_path)
Reporting_Dict = main.Reporting_Dict.copy()#{k:v for k, v in main.Reporting_Dict.items()}
Building_Dimension = [x for x in main.Building_Dimension]


first_print = True
normal_sleep = 2
window_transition = 1
bays_X = Building_Dimension[3]
bays_Y = Building_Dimension[4]
storey_Z = Building_Dimension[5]
elevation_grids = bays_X+1+bays_Y+1
plan_grids = storey_Z+1
critical_grid = 6 #Where Stresses are computed for substuiting in the report. A, B, C, 1, 2 is taken as NO. 1,2,3,4, 5.
timer_a,timer_b,timer_c,timer_d,timer_e = 2, 4, 6, 8, 10
parameters_index = ["LR", "RP", "SR", "PMM", "BC", "CB", "GR", "IP", "IS", "IA"]

general_prints = [] #All other than response prints
response_prints = []

#-------------
#| A | B | C |
#-------------
#| D | E | F |
#-------------
#| G | H | I |
#-------------

#A
def click_topleft_one_third(box):
    x = int(box[0]) + int(box[2])*1/6
    y = int(box[1]) + int(box[3])*1/6
    pyautogui.click(x,y)
# B
def click_top_one_third(box):
    x = int(box[0]) + int(box[2]) / 2
    y = int(box[1]) + int(box[3]) * 1 / 6
    pyautogui.click(x, y)
#C
def click_topright_one_third(box):
    x = int(box[0]) + int(box[2])*5/6
    y = int(box[1]) + int(box[3])*1/6
    pyautogui.click(x,y)
#D
def click_left_one_third(box):
    x = int(box[0]) + int(box[2])*1/6
    y = int(box[1]) + int(box[3]) / 2
    pyautogui.click(x,y)
#F
def click_right_one_third(box):
    x = int(box[0]) + int(box[2])*5/6
    y = int(box[1]) + int(box[3]) / 2
    pyautogui.click(x,y)
#G
def click_bottomleft_one_third(box):
    x = int(box[0]) + int(box[2])*1/6
    y = int(box[1]) + int(box[3])*5/6
    pyautogui.click(x,y)
#H
def click_bottom_one_third(box):
    x = int(box[0]) + int(box[2])/2
    y = int(box[1]) + int(box[3])*5/6
    pyautogui.click(x,y)
#I
def click_bottomright_one_third(box):
    x = int(box[0]) + int(box[2])*5/6
    y = int(box[1]) + int(box[3])*5/6
    pyautogui.click(x,y)

def image_checkpoint_return_location(image, confidence = 0.5):
    try:
        location = pyautogui.locateOnScreen(image, confidence=confidence) #For Confidence to be used in Code pip install opencv-python is needed
        if location is not None:
            return (location)
        else:
            print(f'{image} position is None during runtime')
            #Active value not set as per required by code error
            if image == "longitudinal_reinforcing.png":
                print(f'Set {image} as active before running the code')
            return (location)
    except:
        print(f'{image} is not available in the screen during runtime')
def hold_on(icon):           # If item is present then Code HOLDS
    presence = pyautogui.locateOnScreen(icon)
    while presence != "None":
        sleep(1)
        presence = pyautogui.locateOnScreen(icon)
        if presence == None:
            sleep(1)
            break
def Scroll_Until_Image(Image, Scroll_Number, Scroll_Amount):
    for i in range(Scroll_Number):
        pyautogui.scroll(-Scroll_Amount)
        if pyautogui.locateOnScreen(Image):
            pyautogui.click(pyautogui.center(image_checkpoint_return_location(Image, 0.9)));
            break
    else:
        # If the image is not found after the maximum number of scrolls
        print("Image not found after scrolling.", Image)
def screenshotter(property, dimensions, view = "", index = ''):
    left, top, width, height = dimensions

    hold_on("ok_close_apply.png")
    hold_on("ok.png")


    # Take a screenshot of the specified region
    screenshot = pyautogui.screenshot(region=(left, top, width, height))

    # Save the screenshot as a PDF file
    file_name = os.path.join(main.extracted_data_path, f'{property} {view} {index}.pdf')

    screenshot.save(file_name, 'PDF')

def printer(property,view = "", index = ''):
    if Reporting_Dict["quick_mode"]:
        dimensions = [30, 105, 1325, 580]
        screenshotter(property, dimensions, view=view, index=index)

    else:
        file_name = os.path.join(main.extracted_data_path, f'{property} {view} {index}')
        pyautogui.hotkey('ctrl', 'p');    sleep(timer_c)
        pyautogui.click(image_checkpoint_return_location('print.png'));    sleep(timer_a + 1)
        hold_on("print_hold.png")
        pyautogui.click(image_checkpoint_return_location('file_name.png'));
        pyautogui.write(file_name); pyautogui.press('tab', presses=3); pyautogui.press('enter'); sleep(timer_b)
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("confirm_yesno.png")))
            #PDF Saving Override Condition Check
        except:
            pass
        hold_on("printing_holder.png")
        pyautogui.doubleClick(image_checkpoint_return_location('printer_window.png'));
        pyautogui.hotkey('alt', 'f4'); sleep(timer_a+1)
        general_prints.append(f'{property} {view} {index}.jpg')


def storey_response_printer(property,view = "", index = ''):
    file_name = os.path.join(main.extracted_data_path, f'{property} {view} {index}')

    pyautogui.click(pyautogui.center(image_checkpoint_return_location("storeyresponse_print1.png", 0.9))); sleep(timer_b)
    try:            #Checks the window if maximized then enforces to minimize
        click_top_one_third(image_checkpoint_return_location("response_minimize.png", 0.7)); sleep(window_transition)
    except:
        pass
    pyautogui.click(pyautogui.center(image_checkpoint_return_location("storeyresponse_print2.png", 0.95))); sleep(timer_b)
    pyautogui.click(image_checkpoint_return_location('file_name.png'));
    pyautogui.write(file_name); pyautogui.press('tab', presses=3); pyautogui.press('enter'); sleep(timer_b)
    try:
        pyautogui.click(pyautogui.center(image_checkpoint_return_location("confirm_yesno.png")))
        sleep(timer_e)                           #PDF Saving Override Condition Check
    except:
        sleep(timer_d)
    pyautogui.doubleClick(image_checkpoint_return_location('responseprinter_window.png'));
    pyautogui.hotkey('alt', 'f4'); sleep(timer_a)
    response_prints.append(f'{property} {view} {index}.jpg')

def shortcut_icon(shortcut , icon_image1, confidence=0.9, icon_image2 = ""):
    try:
        if len(shortcut) == 1:
            pyautogui.hotkey(shortcut[0])

        if len(shortcut) == 2:
            pyautogui.hotkey(shortcut[0], shortcut[1])

        if len(shortcut) == 3:
            pyautogui.hotkey(shortcut[0], shortcut[1], shortcut[2])

    except:
        try:
            if icon_image1 == "display_menu.png":
                pyautogui.click(pyautogui.center(image_checkpoint_return_location(icon_image1)));
                sleep(window_transition)
                pyautogui.click(pyautogui.center(image_checkpoint_return_location(icon_image2)));
            else:
                pyautogui.click(pyautogui.center(image_checkpoint_return_location(icon_image1, confidence)))
        except:
            print(f"Error in Hotkey and icon click operation of {icon_image1}")

def try_clicking(image_one, confidence = 0.9):
    try:
        pyautogui.click(pyautogui.center(pyautogui.locateOnScreen(image_one,confidence)))

    except:
        try:
            pyautogui.click(pyautogui.center(pyautogui.locateOnScreen(image_one, confidence)))
        except:
            print("Image not found", image_one)






#___________________________________ETABS PDF Extraction-____________________________________________________________
index_viewdata = {}
def initializer():
    print("Swiftly Switch to Etabs Window under 3 Seconds")
    sleep(timer_c)

def check_for_run():
    pyautogui.press('f5')
    sleep(timer_a)
    timer = 0
    if image_checkpoint_return_location("no_run_cases.png") is None:
        timer = int(input("Tentative Time for analysis of the Structure in Secs?: "))
    else:
        pyautogui.hotkey('enter')
    sleep(timer + 0.5)
def general_images():
    try:
        # 3d Frame Extruded Cover Photo
        shortcut_icon(["alt", "ctrl", "a"], "3d_view.png")
        shortcut_icon(["alt", "ctrl", "b"], "undeformed_shape.png", 0.7)
        pyautogui.hotkey('ctrl', "w");
        sleep(window_transition)
        try_clicking("extrude_frame.png", 0.9)
        try_clicking("extrude_shells.png", 0.9)
        try_clicking("horizon.png", 0.9)

        try:
            pyautogui.click(pyautogui.center(pyautogui.locateOnScreen("extrude_frame.png", 0.9)))
            pyautogui.click(pyautogui.center(pyautogui.locateOnScreen("extrude_shells.png", 0.9)))
        except:
            pass
        pyautogui.press('enter', presses=2);             sleep(timer_c)
        printer("Cover", "3D", 1)


    except:
        print("General Images Function hasnot been Executed")

def stresses_diagrams():
    try:
        shortcut_icon(["alt", "ctrl", "b"], "undeformed_shape.png", 0.7)
        sleep((timer_a - 1)/2);

        pyautogui.hotkey('ctrl', 'shift', 'f2');
        sleep(timer_a - 1);
        pyautogui.press('tab', presses=6);
        pyautogui.press('enter');
        try_clicking("ok_elevation.png", confidence=0.9)
        sleep(timer_a)
        for j in range(0, critical_grid - 1):
            shortcut_icon(["alt", "ctrl", "c"], "up_arrow.png", 0.7)
            sleep(timer_a)

        # Axial Force Diagram
        shortcut_icon(["alt", "ctrl", "e"], "forces_stresses.png", 0.9)
        sleep(timer_a)
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("combo.png", 0.9)));  sleep(timer_a /10)
        except:
            pass
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("axial_tick.png", 0.9)))
        except:
            pass
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("showvalues_tick.png", 0.95)))
        except:
            pass
        sleep(timer_a / 4)
        pyautogui.click(pyautogui.center(image_checkpoint_return_location("ok.png")));
        sleep(timer_a)
        printer("Axial", "Elevation", "1")

        # Moment Diagram
        shortcut_icon(["alt", "ctrl", "e"], "forces_stresses.png", 0.9)
        sleep(timer_a)
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("moment33_tick.png", 0.8)))
        except:
            pass
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("showvalues_tick.png", 0.95)))
        except:
            pass
        sleep(timer_a / 4)
        pyautogui.click(pyautogui.center(image_checkpoint_return_location("ok.png")));
        sleep(timer_a)
        printer("Moment", "Elevation", "1")

        # Shear Force Diagram
        shortcut_icon(["alt", "ctrl", "e"], "forces_stresses.png", 0.9)
        sleep(timer_a)
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("shear22_tick.png", 0.8)))
        except:
            pass
        try:
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("showvalues_tick.png", 0.95)))
        except:
            pass
        sleep(timer_a / 4)
        pyautogui.click(pyautogui.center(image_checkpoint_return_location("ok.png")));
        sleep(timer_a)
        printer("Shear", "Elevation", "1")
    except:
        print("Steess Diagrams Images Function hasnot been Executed")
def forces_diagrams():
    try:
        shortcut_icon(["f7"], "joint_load.png", 0.95)
        # pyautogui.hotkey('ctrl', 'shift', 'f2');
        sleep(timer_a);
        try_clicking('joints_combo.png')

        try_clicking('joints_dropdown.png')
        try_clicking('deads_lives.png')
        try_clicking('deads_lives_checked.png')


        try_clicking("tabulated.png")
        try_clicking("tabulated_checked.png")


        try_clicking('fz.png')
        try_clicking("mx.png")
        try_clicking("my.png")
        try_clicking('ok.png')

        sleep(timer_a)
        pyautogui.hotkey('ctrl', 'shift', 'f1');
        sleep(timer_a)
        try_clicking("base.png")
        sleep(timer_a)
        pyautogui.press('tab', presses=3);
        pyautogui.press('enter');

        sleep(timer_a)
        pyautogui.hotkey('shift', 'f3');         pyautogui.hotkey('shift', 'f3');


        printer("Footing", "Load", "Plan")


    except:
        print("Steess Diagrams Images Function hasnot been Executed")


def design_details():
    # try:
        # Concrete Design Parameters________________________________________________________________________________________
        pyautogui.hotkey('shift', 'ctrl', 'f6')
        sleep(window_transition)
        sleep(2)
        box = image_checkpoint_return_location('design_results_entry.png')
        a = pyautogui.center(box)
        im = Image.open('other_outputs.png')
        width, height = im.size
        unit_height = height / 10
        base_y = a[1] + height + 1.2 * unit_height
        x = a[0]
        y = []
        [y.append(a[1] + unit_height * i) for i in range(1, 11)]
        for i in range(len(main.parameters)):          #xfgdgdfg
            p_index =  parameters_index.index(main.parameters[i])
            pyautogui.hotkey('shift', 'ctrl', 'f6')
            sleep(1)
            pyautogui.click(a[0], a[1])
            pyautogui.click(x, y[p_index])
            pyautogui.hotkey('tab', 'enter')

            if main.parameters[i] in ["LR", "RP", "SR"]:  # For Plan and Elevation data extraction
                # Elevation Screenshot
                sleep((timer_a - 1) / 2);

                pyautogui.hotkey('ctrl', 'shift', 'f2');
                sleep(timer_a - 1)
                pyautogui.press('tab', presses=6);
                pyautogui.press('enter');
                try_clicking("ok_elevation.png", confidence=0.9)
                sleep(timer_a)
                printer(main.parameters[i], "Elevation", 1)
                for j in range(0, elevation_grids - 1):
                    shortcut_icon(["alt", "ctrl", "c"], "up_arrow.png", 0.7)
                    sleep(timer_a)
                    printer(main.parameters[i], "Elevation", j + 2)

                # # Plan Screenshot
                # sleep(timer_b)
                # pyautogui.hotkey('ctrl', 'shift', 'f1');
                # sleep(timer_b)
                # pyautogui.click(pyautogui.center(image_checkpoint_return_location('base.png', 0.9))); sleep(timer_a/10)
                # pyautogui.click(image_checkpoint_return_location('ok.png'));
                #
                # sleep(2)
                # printer(main.parameters[i], "Plan", 1)
                # for j in range(0, plan_grids - 1):
                #     shortcut_icon(["alt", "ctrl", "c"], "up_arrow.png", 0.7)
                #     sleep(timer_a)
                #     printer(main.parameters[i], "Plan", j + 1)

            else:  # For 3D Data Extraction
                shortcut_icon(["alt", "ctrl", "a"], "3d_view.png", 0.9)
                printer(main.parameters[i], "3D", 1)
                pyautogui.hotkey('ctrl', 'shift', 'f3');
                sleep(timer_a - 1)
                pyautogui.click(image_checkpoint_return_location('fast_3d.png'))
                pyautogui.press('tab', presses=8);
                pyautogui.press('enter', presses=36)
                pyautogui.press('tab', presses=3);
                pyautogui.press('enter', presses=16)
                pyautogui.press('tab', presses=3);
                pyautogui.press('enter', presses=6)
                pyautogui.press('tab', presses=5);
                pyautogui.press('enter')
                printer(main.parameters[i], "3D", 2)

    # except:
    #     print("Design Details Function hasnot been Executed")
def storey_response(Code_Used):
    # try:
        # Opening Respose Window

        responses_prints = [ "Displacements", "Drifts"]     #Disp and drifts must be placed sequentially
        dimensions = [474, 126, 890, 575]

        for value in responses_prints:
            shortcut_icon(["alt", "ctrl", "f"], "display_menu.png", 0.7, "storeyresponse_button.png")

            sleep(timer_a / 4)
            # Adjusting Display Parameter
            if value == "Drifts":
                try_clicking("displaytype_response1.png", 0.9)
                sleep(timer_a / 8)

                try:
                    try_clicking("displaytype_response2.png", 0.95)
                except:
                    pyautogui.press('tab');
                    pyautogui.press('enter');
                sleep(timer_a / 8)

                try_clicking("maxdrift_button.png", 0.9)
                sleep(timer_a / 4)




            # Adjusting Load Type Parameter and Calling for Print
            # Along X
            try_clicking("case_response1.png", 0.9)
            sleep(timer_a / 4)
            try:
                try_clicking("case_response2.png", 0.9)
                pyautogui.move(0, 50)
                pyautogui.doubleClick()
            except:
                pyautogui.press('tab');
                pyautogui.press('enter');
                pyautogui.move(0, 50)


            sleep(timer_a / 4)
            if Code_Used == "NBC":
                Scroll_Until_Image("case_eqx_NBC.png", 10, 100)
            else:
                Scroll_Until_Image("case_eqx.png", 10, 100)
            sleep(timer_a / 4)
            if image_checkpoint_return_location("output_type.png", 0.9) is not None:
                pyautogui.click(pyautogui.center(image_checkpoint_return_location("output_type.png", 0.9)));
                pyautogui.click(pyautogui.center(image_checkpoint_return_location("case_response2.png", 0.9)));
                pyautogui.click(pyautogui.center(image_checkpoint_return_location("output_type_max.png", 0.9)));

            if value == "Drifts":
                if Reporting_Dict["story_drift_plot"]:
                    storey_response_printer("Drift", "Elevation", "X")
                if Reporting_Dict["story_drift_plot_ss"]:
                    screenshotter("Drift", dimensions, view = "Elevation SS", index = 'X') #SS - Screenshot

            else:
                if Reporting_Dict["story_deflection_plot"]:
                    storey_response_printer("Displacement", "Elevation", "X")
                if Reporting_Dict["story_deflection_plot_ss"]:
                    screenshotter("Displacement", dimensions, view = "Elevation SS", index = 'X')



            # Along Y
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("case_eqx1.png", 0.7)));
            sleep(timer_a / 10)
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("case_response2.png", 0.95)));
            pyautogui.move(0, 50)
            pyautogui.doubleClick()

            sleep(window_transition)
            if Code_Used == "NBC":
                Scroll_Until_Image("case_eqy_NBC.png", 10, 100)
            else:
                Scroll_Until_Image("case_eqy.png", 10, 100)
            sleep(timer_a / 10)

            if image_checkpoint_return_location("output_type.png", 0.9) is not None:
                pyautogui.click(pyautogui.center(image_checkpoint_return_location("output_type.png", 0.9)));
                pyautogui.click(pyautogui.center(image_checkpoint_return_location("case_response2.png", 0.9)));
                pyautogui.click(pyautogui.center(image_checkpoint_return_location("output_type_max.png", 0.9)));

            if value == "Drifts":
                if Reporting_Dict["story_drift_plot"]:
                    storey_response_printer("Drift", "Elevation", "Y")
                if Reporting_Dict["story_drift_plot_ss"]:
                    screenshotter("Drift", dimensions, view = "Elevation SS", index = 'Y')

            else:
                if Reporting_Dict["story_deflection_plot"]:
                    storey_response_printer("Displacement", "Elevation", "Y")
                if Reporting_Dict["story_deflection_plot_ss"]:
                    screenshotter("Displacement", dimensions, view = "Elevation SS", index = 'Y')

            # Closing the Response Plot Window
            pyautogui.click(pyautogui.center(image_checkpoint_return_location("responsewindow_close.png", 0.9)));
            sleep(timer_a / 4)












    # except:
    #     print("Response Plot Function hasnot been Executed")
#___________________________________Excel PDF Extraction______________________________________________________________
def excel_images(sheet_name, workbook):
    try:
        if sheet_name == "Column Design" and (Reporting_Dict["column_design_excel"] is True):
            image_name = os.path.join(main.extracted_data_path, sheet_name + " Excel.pdf")
        elif sheet_name == "Beam Design" and (Reporting_Dict["beam_design_excel"] is True):
            image_name = os.path.join(main.extracted_data_path, sheet_name + " Excel.pdf")
        elif sheet_name == "Eccentric Footing" and (Reporting_Dict["eccentric_footing_excel"] is True):
            image_name = os.path.join(main.extracted_data_path, sheet_name + " Excel.pdf")
        else:
            image_name = os.path.join(main.extracted_data_path, sheet_name + ".pdf")

        excel_file = client.Dispatch("Excel.Application")
        wb = excel_file.Workbooks.Open(workbook)
        ws = wb.Worksheets(sheet_name)
        ws.ExportAsFixedFormat(0, image_name)
    except:
        print("Either Excel File hasnot been opened Or file isnot inside External files Dir.")

