import main
import os
import docx
import main
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches, Cm, Mm
import win32com.client as win32
import comtypes.client
import xlwings as xw
from time import sleep
from tqdm import tqdm


def reporting_data():
    Reporting_Dict ={}
    Code_Used = ''
    try:
        workbook = xw.Book(main.design_file)

        worksheet = workbook.sheets["Designer Sheet"]
        for shape in worksheet.shapes:
            try:
                name = shape.name
                value = worksheet.api.OLEObjects(name).Object.Value
                Reporting_Dict[name] = value
            except Exception as e:
                print(" Error - ", e)

        if Reporting_Dict['nbc_code']:
            Code_Used = "NBC"
        elif Reporting_Dict['is_code']:
            Code_Used = "IS"
        else:
            Code_Used = "NBC"


        d_parameter_index = worksheet.range("Q7").value
        Building_Length = worksheet.range(f"C{int(d_parameter_index) + 18}").value
        Building_Width = worksheet.range(f"C{int(d_parameter_index) + 19}").value
        Building_Height = worksheet.range(f"C{int(d_parameter_index) + 20}").value
        Bays_X  = int(worksheet.range(f"C{int(d_parameter_index) + 21}").value)
        Bays_Y  = int(worksheet.range(f"C{int(d_parameter_index) + 22}").value)
        Story_Z = int(worksheet.range(f"C{int(d_parameter_index) + 22}").value)
        Bearing_Capacity = int(worksheet.range(f"C{int(d_parameter_index) + 9}").value)
        Plan_Area = int(worksheet.range(f"C{int(d_parameter_index) + 24}").value)

        # workbook.close()

        return(Reporting_Dict,Code_Used, [Building_Length, Building_Width, Building_Height, Bays_X, Bays_Y, Story_Z], Bearing_Capacity,Plan_Area,  worksheet)
    except Exception as e:
        print(" Error - ", e)






main.Reporting_Dict, Code_Used, main.Building_Dimension, main.Bearing_Capacity,main.Plan_Area, worksheet = reporting_data()
Reporting_Dict = main.Reporting_Dict
Building_Dimension = main.Building_Dimension
list_sequence = ["general_images","stresses_diagrams","forces_diagrams","design_details","story_response",
                     "design_excel","etabs_excelsheet","etabs_excel_modifier","etabs_excel_pdf","pdf_to_jpg","data_writer"]
count = 0
Selected_keys = []

import Excel_Word_Handler
import Basic_Tools
import Image_Extractor
import ETABS_Result_Extractor

Image_Extractor.initializer()
def General_Images():
    Image_Extractor.general_images()
def Stress_Diagrams():
    Image_Extractor.stresses_diagrams()
def Forces_Diagrams():
    Image_Extractor.forces_diagrams()
def Design_Details():
    Image_Extractor.design_details()
def Story_Response():
    Image_Extractor.storey_response(Code_Used)
def Design_Excel_Sheet():
    for items in main.excel_sheets:
        if (items == "Column Design" and (Reporting_Dict["column_design_excel"] is False)
            or items == "Beam Design" and (Reporting_Dict["beam_design_excel"] is False)) \
                or items == "Eccentric Footing" and (Reporting_Dict["eccentric_footing_excel"] is False):
            pass
        else:
            Image_Extractor.excel_images(items, main.design_file)
def ETABS_Generated_Excel():
    ETABS_Result_Extractor.model_initializer()
    ETABS_Result_Extractor.GetTableForDisplayArray()
def ETABS_Excel_Processor():
    ETABS_Result_Extractor.Excel_Modifier(Building_Dimension)
def ETABS_Excel_to_PDF():
    for items in main.result_sheets:
        while len(items) > 30:
            items = " ".join(items.split(" ")[0: - 1])
        Image_Extractor.excel_images(items, main.result_file)
def PDFs_to_JPG_Converter():
    os.chdir(main.extracted_data_path)
    Basic_Tools.pdftoimage(Reporting_Dict)
def Data_Writing():
    items_group = Excel_Word_Handler.items_grouper()

    document = docx.Document(main.word_path)
    para = document.paragraphs
    for items in items_group:
        Excel_Word_Handler.placeholder_writer(para, items.split(" ")[0], items.split(" ")[1])
    document.save(main.word_path1)

    Excel_Word_Handler.word_context()
    doc = DocxTemplate(main.word_path1)
    Excel_Word_Handler.image_context(doc)

    context = Excel_Word_Handler.context_word | Excel_Word_Handler.context_image
    doc.render(context)
    doc.save(main.word_path2)




def Main_Function(key):
    # try:
        if key == "general_images" and Reporting_Dict[key]:
            General_Images()
        if key == "stresses_diagrams" and Reporting_Dict[key]:
            Stress_Diagrams()
        if key == "forces_diagrams" and Reporting_Dict[key]:
            Forces_Diagrams()
        if key == "design_details" and Reporting_Dict[key]:
            Design_Details()
        if key == "story_response" and Reporting_Dict[key]:
            Story_Response()
        if key == "design_excel" and Reporting_Dict[key]:
            Design_Excel_Sheet()
        if key == "etabs_excelsheet" and Reporting_Dict[key]:
            ETABS_Generated_Excel()
        if key == "etabs_excel_modifier" and Reporting_Dict[key]:
            ETABS_Excel_Processor()
        if key == "etabs_excel_pdf" and Reporting_Dict[key]:
            ETABS_Excel_to_PDF()
        if key == "pdf_to_jpg" and Reporting_Dict[key]:
            PDFs_to_JPG_Converter()
        if key == "data_writer" and Reporting_Dict[key]:
            Data_Writing()

    # except Exception as e:
    #     print("Error Source: ", key, " Error - ", e)


for key in list_sequence:
    if Reporting_Dict[key]:
        count += 1
        Selected_keys.append(key)


progress_bar = tqdm(total=count)

for i in range(count):
    Main_Function(Selected_keys[i])
    progress_bar.update(1)
    current_progress = progress_bar.n / progress_bar.total
    print(f"Current progress: {current_progress:.0%}")
progress_bar.close()

