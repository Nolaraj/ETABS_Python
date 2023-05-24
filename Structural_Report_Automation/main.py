# import Basic_Tools
import os
# global Reporting_Dict
Reporting_Dict = {}
Building_Dimension = []
Bearing_Capacity = 0
Plan_Area = 0
directory = r"C:\Users\Nolaraj Poudel\Desktop\Python Programming\Report Automation Project\Standard Files Revised 1"
os.chdir(directory)

extracted_data = "Extracted Data"
frame_width = 6         #Word Writitng area dimension In Inches
frame_height = 9
# Available Pameters ["LR", "RP", "SR", "PMM", "BC", "CB", "GR", "IP", "IS", "IA"]
## Select only parameters that are needed. (Note: LR, CB are used in report
parameters = ["LR", "CB"]
excel_sheets = ["Slab Design", "Staircase Design", "Seismic IS", "Seismic NBC", "Isolatedfooting Design", "Eccentric Footing", "Design Summary", "Column Design", "Beam Design"]
result_sheets = ["Modal Participating Mass","Modal Load Participation", "Diaphragm Max Over Avg Drifts",
                 "Centers Of Mass And Rigidity","Modal Periods And Frequencies", "Story Drifts", "Joint Reactions"]
###Result_sheets - Prints Extracted sheets to pdf (Provide Extracted Sheet Names Here)
units = [    "lb_in_F","lb_ft_F","kip_in_F","kip_ft_F","kN_mm_C","kN_m_C","kgf_mm_C","kgf_m_C","N_mm_C","N_m_C ","Ton_mm_C","Ton_m_C ","kN_cm_C ","kgf_cm_C ","N_cm_C ", "Ton_cm_C"]



#For implementation Phase
# excel_path = Basic_Tools.file_path(title="Enter the designer excel path",filetype=("Excel files", ["*.xlsx", "*.xls"]))
# word_path  = Basic_Tools.file_path(title="Enter the designer excel path",filetype=("Word files", ["*.docx", "*.doc"]))


#For developer Phase
# excel_path = os.path.join(os.getcwd(), "External Files", "Design Sheet.xlsx")
word_path = os.path.join(os.getcwd(), "External Files", "Report.docx")
word_path1 = os.path.join(os.getcwd(), "External Files", "Report1.docx")
word_path2 = os.path.join(os.getcwd(), "External Files", "Report2.docx")



# annex_path  = os.path.join(os.getcwd(), "External Files", "Annex.docx")
# rendered_annex_path1  = os.path.join(os.getcwd(), "External Files", "Annex_rendered1.docx")
# rendered_annex_path2  = os.path.join(os.getcwd(), "External Files", "Annex_rendered2.docx")

etabs_icon_path = os.path.join(os.getcwd(), "External Files", "Etabs Icons")
extracted_data_path = os.path.join(os.getcwd(), "External Files", extracted_data)
externalfiles_path = os.path.join(os.getcwd(), "External Files")
design_file  = os.path.join(os.getcwd(), "External Files", "Design File.xlsm")
result_file = os.path.join(extracted_data_path,"Results.xlsx")

if os.path.exists(extracted_data_path) is False:
    os.makedirs(extracted_data_path)

#______________________________________________________________
# all_records = Image_Extractor.general_prints + Image_Extractor.response_prints

#Table Keys for extracting data from the ETABS Database - Provide actual ETABS Keys as in text file
Table_Keys = ["Modal Participating Mass Ratios","Modal Load Participation Ratios", "Diaphragm Max Over Avg Drifts",
              "Centers Of Mass And Rigidity","Modal Periods And Frequencies", "Story Drifts", "Joint Reactions"]

#_______________________________________________
#Report Selection
Reporting_Parameters = ['story_deflection_plot', 'story_drift_plot', 'all_reinforcement_demand', 'column_design_etabs',
                        'column_design_excel', 'beam_design_etabs', 'beam_design_excel', 'footing_loads',
                        'isolated_footing_excel', 'eccentric_footing_excel', 'modal_mass_excel', 'diaphragm_max_avg'
                        , 'nbc_code', 'is_code', 'nbc_is_combined', 'story_deflection_plot_ss', 'story_drift_plot_ss']