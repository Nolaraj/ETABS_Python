from tkinter import Tk, filedialog
import main
import os
import cv2
import pdf2image
Reporting_Dict = main.Reporting_Dict


#Extract Folder path from the dialog box
def folder_path(title="Enter the file path"):
    root = Tk()             # pointing root to Tk() to use it as Tk() in program.
    root.withdraw()         # Hides small tkinter window.
    root.title()
    root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.
    # root.attributes()
    open_file = filedialog.askdirectory(title = title) #,filetypes = (("Excel files", ".xlsx .xls"),))   # Returns opened path as str
    return open_file

#Extract selected file path from the dialog box
def file_path(title="Enter the file path",filetype=("Excel files", ["*.xlsx", "*.xls"])):
    root = Tk()             # pointing root to Tk() to use it as Tk() in program.
    root.withdraw()         # Hides small tkinter window.
    root.title()
    root.attributes('-topmost', True) # Opened windows will be active. above all windows despite of selection.
    # root.attributes()
    open_file = filedialog.askopenfilename(title = title,filetypes = ((filetype[0], filetype[1][0]), (filetype[0], filetype[1][1])))   # Returns opened path as str
    return open_file

#______________________________________________________Data Processing_________________________________________________
def crop_optimizer(image, long_start, long_end, trans_start, trans_end, long_interval, trans_interval, dir = "X"):   #Until non white dot (from 4 sides)
    sub_interval = int(-(long_interval/25))
    for i in range(long_start, long_end, long_interval):
        for j in range(trans_start, trans_end, trans_interval*2):
            if dir == "X":
                x = i; y=j
            if dir == "Y":
                x = j; y=i
            # print(x, y)
            colour = image[y][x]
            if colour[0] < 250 or (colour[1] < 250 or colour[2] < 250):

                for k in range(i, long_start, sub_interval):
                    check = True
                    for l in range(trans_start, trans_end, trans_interval*2):
                        # print(k, l)
                        if dir == "X":
                            x = k;      y = l
                        else:
                            x = l;      y = k
                        colour = image[y][x]
                        # print(colour, i, j, x, y)
                        if colour[0] < 250 or (colour[1] < 250 or colour[2] < 250):
                            check = False

                    if check == True:
                        if (k + sub_interval) > 0:
                            return (k + sub_interval)
                        else:
                            return (0)
                    # print(abs(abs(k) - abs(long_start)))

                    if abs(abs(k) - abs(long_start)) <= abs(sub_interval):
                        # print(abs(abs(k) - abs(long_start)))
                        return k


def pdftoimage(Reporting_Dict):
    # try:
        # Importing pdf and saving to Image Jpeg #os.chdir must be changed to the preferred else (I/O Error: Couldn't open file)
        pdf_directory = main.extracted_data_path
        for items in os.listdir(pdf_directory):
            if items.endswith(".pdf"):
                image_name = f'{items[:-4]}.jpg'
                if image_name.split(" ")[0] in ["Drift", "Displacement"]:  # Processing for Response Prints PDF Data
                    pages = pdf2image.convert_from_path(items)
                    for i in range(len(pages)):
                        pages[i].save(image_name, 'JPEG')

                        img = cv2.imread(image_name)  # Cropping Image (Pre Cropping Preparation)

                        if (Reporting_Dict["story_deflection_plot_ss"] == False) and (Reporting_Dict["story_drift_plot_ss"] == False):
                            cropped_image = img[200:1000, 200:1550]         # H*B = 800*1350
                            cv2.imwrite(image_name, cropped_image)

                        img = cv2.imread(image_name)  # Cropping Image Until non white pixel comes
                        width, height = img.shape[1], img.shape[0]
                        cropped_image = img[crop_optimizer(img, 0, height - 5, 0, width-5, 50, 5, "Y"):
                                        crop_optimizer(img, height - 5, 0, 0, width-5, -50, 5, "Y"),
                                        crop_optimizer(img, 0, width - 5, 50, height - 5, 50, 2, "X"):
                                        crop_optimizer(img, width-5, 0, 50, height-5, -50, 2, "X")]
                        cv2.imwrite(image_name, cropped_image)

                elif image_name.split(" ")[0] in ["Column", "Beam", "Shearwall"]:
                    pages = pdf2image.convert_from_path(items)
                    for i in range(len(pages)):
                        image_name = f'{items[:-4]} {i + 1}.jpg'
                        pages[i].save(image_name, 'JPEG')

                        img = cv2.imread(image_name)
                        cropped_image = img[150:2050, 200:1550]          # H*B = 1900*1350
                        cv2.imwrite(image_name, cropped_image)
                        #
                        img = cv2.imread(image_name)
                        width, height = img.shape[1], img.shape[0]
                        cropped_image = img[crop_optimizer(img, 0, height - 5, 0, width-5, 50, 5, "Y"):
                                        crop_optimizer(img, height - 5, 0, 0, width-5, -50, 5, "Y"),
                                        crop_optimizer(img, 0, width - 5, 50, height - 5, 50, 2, "X"):
                                        crop_optimizer(img, width-5, 0, 50, height-5, -50, 2, "X")]
                        cv2.imwrite(image_name, cropped_image)


                elif image_name.split(" ")[0] in main.parameters + ["Cover", "Axial", "Shear",
                                                                    "Moment", "Footing"]:  # All from Printer function
                    pages = pdf2image.convert_from_path(items)
                    for i in range(len(pages)):
                        pages[i].save(image_name, 'JPEG')


                        if Reporting_Dict["graphics_mode"]:
                            img = cv2.imread(image_name)
                            cropped_image = img[150:2050, 200:1550]          # H*B = 1900*1350
                            cv2.imwrite(image_name, cropped_image)
                        #
                        img = cv2.imread(image_name)
                        width, height = img.shape[1], img.shape[0]

                        cropped_image = img[crop_optimizer(img, 0, height - 5, 0, width-5, 50, 5, "Y"):
                                        crop_optimizer(img, height - 5, 0, 0, width-5, -50, 5, "Y"),
                                        crop_optimizer(img, 0, width - 5, 50, height - 5, 50, 2, "X"):
                                        crop_optimizer(img, width-5, 0, 50, height-5, -50, 2, "X")]
                        cv2.imwrite(image_name, cropped_image)

                elif image_name.split(" ")[0] in ["Slab", "Staircase", "Isolatedfooting", "Eccentric", "Seismic", "Modal", "Story", "Centers", "Diaphragm", "Design"]\
                        or (image_name.split(" ")[1] in ["Summary"]\
                        or image_name.split(" ")[-1] in ["Excel.pdf"]):
                    pages = pdf2image.convert_from_path(items)
                    for i in range(len(pages)):
                        image_name = f'{items[:-4]} {i + 1}.jpg'
                        pages[i].save(image_name, 'JPEG')

                        img = cv2.imread(image_name)
                        width, height = img.shape[1], img.shape[0]

                        cropped_image = img[crop_optimizer(img, 0, height - 50, 0, width - 50, 50, 5, "Y"):
                                            crop_optimizer(img, height - 50, 0, 0, width - 50, -50, 5, "Y"),
                                        crop_optimizer(img, 0, width - 50, 50, height - 50, 50, 2, "X"):
                                        crop_optimizer(img, width - 50, 0, 50, height - 50, -50, 2, "X")]
                        cv2.imwrite(image_name, cropped_image)

                else:  # Processing for General Prints PDF Data and OTHER ELSE All
                    pages = pdf2image.convert_from_path(items)
                    for i in range(len(pages)):
                        image_name = f'{items[:-4]} {i + 1}.jpg'
                        pages[i].save(image_name, 'JPEG')

    # except:
    #     print("Error in Image Editor")

def txt_writer():
    try:
        file1 = open("Records.txt", "w")
        writer = []
        for i in range(len(main.all_records)):
            writer.append(str(main.all_records[i]));
            writer.append("\n")
        file1.writelines(writer)
        file1.close()
    except:
        print("Error in Text Writer")




