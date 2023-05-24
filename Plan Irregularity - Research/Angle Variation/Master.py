import Post_Processing
import User_Input
from User_Input import *
from Pre_Processings import *
from ETABS import *
from Post_Processing import *

def ETABS_Caller():
    if Critical_Angle_Determination:
        global Calculation_Option
        global max_Angle

        if Calculation_Option == 4:
            for nth_iteration in range(1, len(angles) + 1):
                for step in range(1, 4):
                    max_Angle = 0
                    Calculation_Option = step
                    for i in range(0, 2):
                        Definition_Center(i, angles[nth_iteration - 1])
                        ModellingAndResults_Center(nth_iteration, angles, i)
        else:
            for nth_iteration in range(1, len(angles) + 1):
                max_Angle = 0
                for i in range(0, 2):
                    Definition_Center(i, angles[nth_iteration - 1])
                    ModellingAndResults_Center(nth_iteration, angles, i)



    else:
        for nth_iteration in range(1, len(angles) + 1):
            Definition_Center(angle=angles[nth_iteration - 1])
            ModellingAndResults_Center(nth_iteration, angles)

    Response_correction = False
    if Response_correction:
        if len(Earthquake_Casesx) > 0:
            looper = len(Earthquake_Casesx)
        elif len(Earthquake_Casesy) > 0:
            looper = len(Earthquake_Casesy)
        else:
            print("No Set of ESM Load case found for correction")

        # try:
        for i in range(0, looper):
            for k in range(1, len(angles) + 1):
                # X SF Correction
                earthquake_casex = Earthquake_Casesx[i]
                response_casex = Response_Casesx[i]
                f_result = "Result.xlsx"
                wb = load_workbook(f_result, data_only=True)
                try:
                    ws = wb["BaseReactWithCentroid"]
                    indexer(wb, "BaseReactWithCentroid", "FX")

                except:
                    ws = wb["Base Reactions"]
                    indexer(wb, "Base Reactions", "FX")

                angle = angles[k - 1]
                earthquake_valuex = 0
                response_valuex = 0

                for l in range(block_start_row_i[k - 1], block_end_row_i[k - 1]):

                    if ws.cell(row=l, column=Pre_Processings.load_case_col).value == earthquake_casex:
                        earthquake_valuex = float(
                            ws.cell(row=l, column=Pre_Processings.start_column_i).value)
                    if ws.cell(row=l, column=Pre_Processings.load_case_col).value == response_casex:
                        response_valuex = float(
                            ws.cell(row=l, column=Pre_Processings.start_column_i).value)

                if earthquake_valuex != 0 and response_valuex != 0:
                    Resp_Spectrum_SFx = [abs((earthquake_valuex / response_valuex) * original_RS_SFx)]
                else:
                    Resp_Spectrum_SFx = original_RS_SFx

                # Y SF Correction
                earthquake_casey = Earthquake_Casesy[i]
                response_casey = Response_Casesy[i]
                earthquake_valuey = 0
                response_valuey = 0
                try:
                    indexer(wb, "BaseReactWithCentroid", "FY")

                except:
                    indexer(wb, "Base Reactions", "FY")
                for l in range(block_start_row_i[k - 1], block_end_row_i[k - 1]):
                    if ws.cell(row=l, column=Pre_Processings.load_case_col).value == earthquake_casey:
                        earthquake_valuey = float(
                            ws.cell(row=l, column=Pre_Processings.start_column_i).value)
                    if ws.cell(row=l, column=Pre_Processings.load_case_col).value == response_casey:
                        response_valuey = float(
                            ws.cell(row=l, column=Pre_Processings.start_column_i).value)

                if earthquake_valuey != 0 and response_valuey != 0:
                    Resp_Spectrum_SFy = [abs((earthquake_valuey / response_valuey) * original_RS_SFy)]
                else:
                    Resp_Spectrum_SFy = original_RS_SFy

                wb.close()

                print(earthquake_valuex, response_valuex, Resp_Spectrum_SFx, "Resp Spec", k)
                print(earthquake_valuey, response_valuey, Resp_Spectrum_SFy, "Resp Spec", k)

                # # #Calling Zone
                f_result = f'Result {earthquake_casex}-{response_casex}.xlsx'
                Definition_Center()
                ModellingAndResults_Center(k, angles)
                print(f"Angle {angles[k - 1]} writing has successfully been completed, Thank You")
        # except:
        #     print("TRY Statement hasnot been executed due to internal error, Check Either CASEs, Scale Factor Provided, Deviation Angles")



def AnalysisResultProcessing_Caller():
    global key
    for key, value in sheet_title.items():
        if len(value) > 0:
            for title in value:
                control_center_Analysis_Results(key, title)
                sheet_name = key

def TabledOutputProcessing_Caller():
    def General_processor():
        control_center_Table_Results()

    def JointsElements_processor():
        # _______________________________Story Displacements and Drifts - Tabled Results Extractor_________________
        bulk_numbers = 4        #Number of joints considered in a floor ie extracted from Output_joints list
        for i in range(0, Floors_number):
            wb1 = Post_Processing.wb1
            output_sheet = f"Output {len(wb1.sheetnames) - 1}"

            for k in range(2,3):
                OutputCase = []
                Unique_Names = []
                OutputCase.append(User_Input.Output_cases[k])
                OutputCase.append(None)

                for j in range(0, bulk_numbers):
                    Unique_Names.append(Output_joints[i * bulk_numbers + j])
                Unique_Names.append(None)

                sorting_data = {
                    "Joint Displacements": [["OutputCase", OutputCase,
                                             ["Story", "OutputCase", "UniqueName", "Ux", "Uy", "Uz", "Rx", "Ry",
                                              "Rz"],
                                             ["Ux", "Uy", "Uz", "Rx", "Ry", "Rz"]]]
                        # ,
                        #                     ["UniqueName", Unique_Names]
                        #                     ]
                    # #                         ,
                    # "Joint Drifts": [["OutputCase", OutputCase,
                    #                   ["Story", "OutputCase", "UniqueName", "Dispx", "DispY", "DriftX", "DriftY"],
                    #                   ["Dispx", "DispY", "DriftX", "DriftY"]],
                    #                  ["UniqueName", Unique_Names]
                    #                  ]
                }
                User_Input.sort_data = sorting_data
                control_center_Table_Results()


    if JointsElements_Output:
        JointsElements_processor()
    else:
        General_processor()




#
ETABS_Caller()
Processing_Initializor()
# AnalysisResultProcessing_Caller()
# TabledOutputProcessing_Caller()




#
#
#
# #
# for sheet_names in CA_SheetNames:
#     if CA_SheetNames.index(sheet_names) == 0:
#         Post_Processing.Customized_Arranger(sheet_names)