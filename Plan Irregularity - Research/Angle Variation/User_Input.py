Titles = ["P", "V2", "V3", "T", "M2", "M3"] #Donot modify
elements = ["213313","221311"] #One for the Absoulute Value and One for percentage difference type Column group
Output_joints = ["1", "11", "71","61","2", "12", "72", "62", "73", "78", "108", "103","109", "114","144","139","145","150","180","175","181","186","216","211","217","222","252","247","253","258","288","283"]
Output_cases = ["Response Spectrum X", "Response Spectrum Y", "Response Spectrum Critical"]
data_column = {"S.N.": 1, "Element": 2, "Combo": 3, "Angle": 4, "Parameter1": 5,
               "Start": 5, "SPD": 6, "Centre": 7, "CPD": 8, "End": 9, "EPD": 10, "8": "P",
               "9": "V2", "10": "V3", "11": "T", "12": "M2", "13": "M3"}  #
data_column_BaseReaction = {"S.N.": 1, "Combo": 2, "Angle": 3, "Number of Storey": 4, "Result's Point Location":5, "Fx": 6,
               "Fy": 7, "Fz": 8, "Mx": 9, "My": 10, "Mz": 11}
properties_ = ["Fx", "Fy", "Fz", "Mx", "My", "Mz"]
########### Extract the required data from here ["P", "V2", "V3", "T", "M2", "M3"] 8,9,10,11,12,13
create_chart = True
parameter = [0]  #Provide only one value inside it in this phase

CoOrdinatesFrom_Excel = False
#IF False Specify the type of co-Ordinates (1, 2, 3...)
CoOrdinates_Type = 3
Quadrilateral_Boundary = True #False will set the beams in unconservative and resulting in more than 4 beams from a joint


#Response Spectrum X and Y Modifiers
#Response Spectrum [Name, NumberLoads = 1,LoadName = ["U1"],Func = ['IS 1893-2016'],SF = [5.6](in m/s^2),CSys = ['Global'],Ang = [0.0]]
Resp_Spectrumx = ["Response Spectrum X", 1, ["U1"], ['IS 1893-2016'], [9806.65], ['Global'], [0.0]]     #Scale factor is Dummy and isnot used until corrected by multiplier in subsequent step
Resp_Spectrumy = ["Response Spectrum Y", 1, ["U2"], ['IS 1893-2016'], [9806.65], ['Global'], [0.0]]     #Scale factor is Dummy and isnot used until corrected by multiplier in subsequent step
Resp_SpectrumCric = ["Response Spectrum Critical", 1, ["U2"], ['IS 1893-2016'], [9806.65], ['Global'], [0.0]]     #Scale factor and Loading Angle is Dummy and isnot used until corrected by multiplier in subsequent step
Diaphragm_eccentricity = [["Response Spectrum X", 0.05],["Response Spectrum Y", 0.05], ["Response Spectrum Critical", 0.05]]
Diaphragm_Ecc_for_other = 0.02      # Other than in above mentioned line

Default_RS_SF = 9806.65
RS_X_Y_Multiplier = 1.90               #Overrides above mentioned Response Spectrum SF along X and Y by multiplying with default
RS_Critical_Multiplier = 1.90         #Overrides Response Spectrum SF along Critical direction (angle of Model as provided in CriticalAngles_DA_RS)

Resp_Spectrum_SFx = [Default_RS_SF * RS_X_Y_Multiplier]      #This value is used for initial modelling  proccess
Resp_Spectrum_SFy = [Default_RS_SF * RS_X_Y_Multiplier]      #This value is used for initial modelling  proccess
Resp_Spectrum_SF_Cr = [Default_RS_SF * RS_Critical_Multiplier]      #This value is used for initial modelling  proccess

CriticalAngles_DA_RS = {0: 0, 5: 0.215, 10: 0.513, 20: 1.78, 40: 6.89, 60: 8.16, 30: 4.253, 50: 8.341, 70: 6.636, 80: 4.26}            #Value in specific index corresponds to the Corresponding Response Spectrum Load Angle For same indexed deviation_angles Model


#Response Scale Factor Correction
Response_correction = True           #Corrects Response spectrum analysis with Scale factor of EQ/RS for base shear
Earthquake_Casesx = ["EQx"]                 #(All have same num of item inside list) #Only used to trap Shear Force Value and modify Resp_Spectrumx List's Name's SF
Response_Casesx = ["Response Spectrum X"]       #(All have same num of item inside list) #Only used to trap Shear Force Value and modify Resp_Spectrumx List's Name's SF
Earthquake_Casesy = ["EQy"]                 #(All have same num of item inside list) #Only used to trap Shear Force Value and modify Resp_Spectrumy List's Name's SF
Response_Casesy = ["Response Spectrum Y"]       #(All have same num of item inside list) #Only used to trap Shear Force Value and modify Resp_Spectrumy List's Name's SF

#Critical Angle Determination Response Spectrum Cases
Critical_Angle_Determination = False         #Predefined Model Should be made False to activate this step
critical_angle_UpperLimit = 90
loads_applied = ["RS", 1, ["U1"], ['IS 1893-2016'], [9806.65], ['Global'], [0.0]]
Diaphragm_Ecc = 0.05
critical_cases = []
max_Angle = 0
Calculation_Option = 4    #Set 1 to find cric. angle when maximum Reaction is obtained in refeence direction 2. it scans for minimum reaction in transverse dir 3. it scans for critical angle if the Reaction between Reference Dir and Transverse is Maximum 4. To perform all mentioned 3 steps


#### Creating Cases for the Results
cases_output = ["Modal", "Dead", "Live", "EQx ULS", "EQx SLS", "EQy ULS", "EQy SLS", "EQx", "EQy", "Response Spectrum"]
# Provide item named 'Response Spectrum' only to capture all RS cases and critical angle determination else provide respective RS names to abort mentioned activities
#, "Response Spectrum XY", "Response Spectrum X", "Response Spectrum + X","Response Spectrum - X", "Response Spectrum Y",  "Response Spectrum + Y", "Response Spectrum - Y", "Response Spectrum 45"]
combos_output = ["Seismic Weight"]

## When predefined (Materials, Sections, Patterns, Case, Combination) Etabs file is used then set following as True
predefined_model = False

## Over-ride for Angle variation to be used than  Excel Provided (If Excel data to be used make it as empty list)
deviation_angles = [50]#, 5, 10, 20, 40, 60]

#____________________________________________________ANALYSIS RESULTS TYPE DATA____________________________________
#Results Refined Data
refined_firstcols = ["Angle Deviation", "Obj", "LoadCase"]
sheet_title = {
        "BaseReact"                    : [] ,   #["NumberResults", "LoadCase", "StepType", "StepNum", "Fx", "Fy", "Fz", "Mx", "ParamMy", "Mz", "gx", "gy", "gz"]
        "StoryDrifts"                  : [] ,   #["NumberResults", "Story", "LoadCase", "StepType", "StepNum", "Direction", "Drift", "Label", "X", "Y", "Z"]
        "BaseReactWithCentroid"        : ["FX", "FY"],#"FX", "FY", "FZ", "MX", "ParamMy", "MZ"] ,   #["NumberResults", "LoadCase", "StepType", "StepNum", "FX", "FY", "FZ", "MX", "ParamMy", "MZ", "GX", "GY", "GZ", "XCentroidForFX", "YCentroidForFX", "ZCentroidForFX", "XCentroidForFY", "YCentroidForFY", "ZCentroidForFY", "XCentroidForFZ", "YCentroidForFZ", "ZCentroidForFZ"]
        "BucklingFactor"               : [] ,   #["NumberResults",        "LoadCase",        "StepType",        "StepNum",        "Factor"]
        "FrameForce"                   : [] ,   #["NumberResults", "Obj", "ObjSta", "Elm", "ElmSta", "LoadCase", "StepType", "StepNum", "P", "V2", "V3", "T", "M2", "M3"]
        "FrameJointForce"              : [] ,   #["NumberResults", "Obj", "Elm", "PointElm", "LoadCase", "StepType", "StepNum", "F1", "F2", "F3", "M1", "M2", "M3"]
        "JointAcc"                     : [] ,   #["NumberResults", "Obj", "Elm", "LoadCase", "StepType", "StepNum", "U1", "U2", "U3", "R1", "R2", "R3"]
        "JointAccAbs"                  : [] ,   #["NumberResults", "Obj", "Elm", "LoadCase", "StepType", "StepNum", "U1", "U2", "U3", "R1", "R2", "R3"]
        "JointDrifts"                  : [] ,   #["NumberResults", "Story", "Label", "Name", "LoadCase", "StepType", "StepNum", "DisplacementX", "DisplacementY", "DriftX", "DriftY"]
        "JointReact"                   : [] ,   #["NumberResults", "Obj", "Elm", "LoadCase", "StepType", "StepNum", "F1", "F2", "F3", "M1", "M2", "M3"]
        "JointVel"                     : [] ,   #["NumberResults",  "Obj",  "Elm",  "LoadCase",  "StepType",  "StepNum",  "U1",  "U2",  "U3",  "R1",  "R2",  "R3"]
        "JointVelAbs"                  : [] ,   # ["NumberResults", "Obj", "Elm", "LoadCase", "StepType", "StepNum", "U1", "U2", "U3", "R1", "R2", "R3"]
        "ModalLoadParticipationRatios" : [] ,   #["NumberResults", "LoadCase", "ItemType", "Item", "Stat", "Dyn"]
        "ModalParticipatingMassRatios" : [],    #"Period", "UX", "UY", "SumUX", "SumUY"] ,     #["NumberResults", "LoadCase", "StepType", "StepNum", "Period", "UX", "UY", "UZ", "SumUX", "SumUY", "SumUZ", "RX", "RY", "RZ", "SumRX", "SumRY", "SumRZ"]
        "ModalPeriod"                  : [],    #"Period", "Frequency", "CircFreq", "EigenValue"] ,   #["NumberResults", "LoadCase", "StepType", "StepNum", "Period", "Frequency", "CircFreq", "EigenValue"]
        "ModeShape"                    : [],    #"U1", "U2", "U3", "R1", "R2", "R3"] ,   #["NumberResults", "Obj", "Elm", "LoadCase", "StepType", "StepNum", "U1", "U2", "U3", "R1", "R2", "R3"]
        "SectionCutAnalysis"           : [] ,   #["NumberResults", "SCut", "LoadCase", "StepType", "StepNum", "F1", "F2", "F3", "M1", "M2", "M3"]
        "SectionCutDesign"             : [] ,   #["NumberResults", "SCut", "LoadCase", "StepType", "StepNum", "P", "V2", "V3", "T", "M2", "M3"]
        "JointDispl"                   : [] ,   #["NumberResults", "Obj","Elm","LoadCase","StepType","StepNum","U1","U2","U3","R1","R2","R3"]
        "JointDisplAbs"                : [] ,   #["NumberResults", "Obj", "Elm", "LoadCase", "StepType", "StepNum", "U1", "U2", "U3", "R1", "R2", "R3"]
        "AssembledJointMass"           : [] ,   #["NumberResults", "PointElm", "U1", "U2", "U3", "R1", "R2", "R3"]
                     }

#Items enforced to create chart with y_values aligned horizontally
horizontally_valued = ["ModalParticipatingMassRatios"]


#Assembled data Pretitles ie done for row above title in arranged data
assembled_pretitles = [ "StepNum", "Direction", "Story"]

#Customized Arrangeer Sheetnames (Sheet names also can be provided manually by erasing loop in Mastaer (except calling line) and giving sheet details in function)
CA_SheetNames = ["Disp RS X", "Disp RS Y", "Disp RS Critical", "Drift RS X", "Drift RS Y", "Drift RS Critical"]

#Customized Arranger Chart List
CA_SeriesColour = ["D30000", "0018F9", "3BB143", "FCE205", "FC0FC0", "2B1700", "000080", "FF00FF", "4CBB17", "C21807", "0080FE"]
CA_Markers = [ "x", "triangle", "square",  "circle",  "plus", "diamond", "star", "dash", "dot",  "picture",  "auto"]
CA_Dashes = ["solid", "sysDash", "dash", "dot", "sysDot", "lgDashDot",  "dashDot", "lgDashDotDot", "sysDashDotDot", "lgDash", "sysDashDot"]

#_________________________TABLED DATA EXTRACTION_________________________________________________________________
Floors_number = 8    #Used for disp and drift extractor It is floors numbers ie provide as: stories + 1
JointsElements_Output = False         #Used for heavy type of extraction if True, If false General method of extraction is enforced
OCase_UniqueNameSet_Use  = True       #Used for fast processing in case of Heavy data to filter Unique names in set in inital process
#JointsElements_Output and OCase_UniqueNameSet_Use both set as true causes escalating rise in Computation SPeed
# Set_titles = ["OutputCase", "UniqueName"]
#Table Sorting Data (Provide Keys as Provided in Table Keys Only)
x_title = "Angle Deviation"         #Used in charts
series_title = "OutputCase"


#Table Keys for extracting data from the ETABS Database
Table_Keys = ["Joint Design Reactions",
              "Joint Displacements",
              "Joint Drifts"]

# key = Sheet_Name,   Value (First Item) = Title,   Value (Second Item (LIST)) = Data to be sorted
# Value(Third) = Col that needs to be extracted in Refined  Forth Value = Titles for charts creation
# Provide all in string form


sort_data = {"Joint Displacements": [["OutputCase", ["Response Spectrum Critical"], ["Story", "OutputCase", "UniqueName",  "Ux",	"Uy",	"Uz",	"Rx",	"Ry",	"Rz"], ["Ux",	"Uy",	"Uz",	"Rx",	"Ry",	"Rz"]]]

             } #,
            #                      ["UniqueName", ["253","258","288","283", None]]
            #                      ],
            #
            # "Joint Drifts": [["OutputCase", ["Response Spectrum Y"], ["Story", "OutputCase", "UniqueName",  "Dispx",	"DispY",	"DriftX",	"DriftY"], ["Dispx",	"DispY",	"DriftX",	"DriftY"]],
            #                      ["UniqueName", ["253", "258", "288", "283", None]]
            #                      ]
            #
            #   }
