from connect import *
import connect, ctypes, sys

script_version = 1.00
script_name = 'Breast VMAT ROBUST'
#written by Ian Gleeson
#runs in CPython 2.7 64bit

#get current case, patient and examination
case = get_current("Case")
patient = get_current("Patient")
examination = get_current("Examination")
plan = get_current ("Plan")

#gives error if not HFS and then exits when user hits ok. runs with C python but text errors with iron python cant see all letters
if examination.PatientPosition != 'HFS':
    rtn = ctypes.windll.user32.MessageBoxA(0, 'Script only for HFS!', 'Breast VMAT ROBUST v' + str(script_version), 16)
    if rtn == 1:
        exit()

#checks that plan name must be R_Breast_Initial or L_Breast_Initial and exits if we dont have it giving a messege
plan = case.TreatmentPlans[plan.Name]

while plan.Name not in ['L_Breast_Initial', 'R_Breast_Initial']:
    rtn = ctypes.windll.user32.MessageBoxA(0, "Plan Name must be called 'L_Breast_Initial' or 'R_Breast_Initial'!", 'Breast VMAT ROBUST v' + str(script_version), 16)
    if rtn == 1:
        exit()
    else:
        pass

#any plan in case cannot be called L_Breast_VMAT or R_Breast_VMAT or LBreast_VMATSIB or RBreast_VMATSIB as thats what the copy plan can be called later
for p in case.TreatmentPlans:
    if p.Name in ['L_Breast_VMAT', 'R_Breast_VMAT', 'LBreast_VMATSIB', 'RBreast_VMATSIB']:
        rtn = ctypes.windll.user32.MessageBoxA(0, "No Plan Names in Case can be called 'L_Breast_VMAT' or 'R_Breast_VMAT' or 'LBreast_VMATSIB' or 'RBreast_VMATSIB'!", 'Breast VMAT ROBUST v' + str(script_version), 16)
        if rtn == 1:
            exit()
        else:
            pass
            

#checks if we have an roi called CTVp_4000 and exits if we dont giving a messege
rois = case.PatientModel.RegionsOfInterest 

roi_list = []
    
for roi in rois:
    roi_list.append(roi.Name)
    

if 'CTVp_4000' not in roi_list:
    rtn = ctypes.windll.user32.MessageBoxA(0, "No ROI called 'CTVp_4000' found!", 'Breast VMAT ROBUST v' + str(script_version), 16)
    if rtn == 1:
        exit()
    else:
        pass

#checks if Ipsi_Lung there. To be changed to ipsi lung soon
if 'Ipsi_Lung' not in roi_list:
    rtn = ctypes.windll.user32.MessageBoxA(0, "No ROI called 'Ipsi_Lung' found!", 'Breast VMAT ROBUST v' + str(script_version), 16)
    if rtn == 1:
        exit()
    else:
        pass


#checks have roi called Skin as needed for fusions and flash?
if 'Skin' not in roi_list:
    rtn = ctypes.windll.user32.MessageBoxA(0, "No ROI called 'Skin' found!", 'Breast VMAT ROBUST v' + str(script_version), 16)
    if rtn == 1:
        exit()
    else:
        pass

#checks not same organ simulation motion Examination Group present and errors
for ExaminationGroups in case.ExaminationGroups:
    if ExaminationGroups.Name=="Simulated organ motion":
        rtn = ctypes.windll.user32.MessageBoxA(0, "Another ExaminationGroup exists with name 'Simulated organ motion'.Please delete Group or rename'!", 'Breast VMAT ROBUST v' + str(script_version), 16)
        if rtn == 1:
            exit()
        else:
            pass

#gets center roi CTVp_4000 coordinates and creates poi in center

ctv_name = 'CTVp_4000'
poi_name = "center_CTVp_4000"

# Get ROI geometries 
roi_geometries = case.PatientModel.StructureSets[examination.Name].RoiGeometries

#get the center of the CTV ROI
try:
    ctv_center = roi_geometries[ctv_name].GetCenterOfRoi()
except:
    print 'Cannot access center of ROI {0}. Exiting script.'.format(ctv_name)
    sys.exit()
    
#save the coordinates in a dictionary
center_CTVp_4000_coordinates = {'x':ctv_center.x, 'y':ctv_center.y, 'z':ctv_center.z}

#delete POI if it already exists
pois = case.PatientModel.PointsOfInterest
try:
    pois[poi_name].DeleteRoi()
    print 'deleted previous POI'
except:
    print 'creating POI'


#Create a POI from the coordinates 
with CompositeAction('Create center_CTVp_4000 POI'):
    center_CTVp_4000_poi = case.PatientModel.CreatePoi(Examination=examination, 
    Point=center_CTVp_4000_coordinates, 
    Name=poi_name, Color="Green", VisualizationDiameter=1, 
    Type="Undefined")



#name of the POI
poi_name = 'center_CTVp_4000'

#access the POI
poi_geometries = case.PatientModel.StructureSets[examination.Name].PoiGeometries

try:
    center_CTVp_4000_poi = poi_geometries[poi_name]
except:
    print 'Error: POI not found' 
    sys.exit()

#save the values of the coordinates
x = center_CTVp_4000_poi.Point.x
y = center_CTVp_4000_poi.Point.y
z = center_CTVp_4000_poi.Point.z

#edit plan name if not called 'L_Breast_Initial' or 'R_Breast_Initial', is they are called correctly then ignore
#looks at ctvp laterality and checks correct name applied depending of center of ctvp right or left to set up point

plan = case.TreatmentPlans[plan.Name]
if x > 0 and plan.Name == 'L_Breast_Initial':
    pass
elif x > 0 and plan.Name != 'L_Breast_Initial':
    plan.Name = 'L_Breast_Initial'
if x < 0 and plan.Name == 'R_Breast_Initial':
    pass
elif x < 0 and plan.Name != 'R_Breast_Initial':
    plan.Name = 'R_Breast_Initial'    

#EDIT BEAMSET name to match plan name
plan.BeamSets[0].DicomPlanLabel = plan.Name

if plan.Name == 'L_Breast_Initial':
    NewPlan_Name = 'L_Breast_VMAT'
elif plan.Name == 'R_Breast_Initial':
    NewPlan_Name = 'R_Breast_VMAT'
    
#copy plan and rename it to new plan name and beamset and set as current
case.CopyPlan(PlanName=plan.Name, NewPlanName=NewPlan_Name, KeepBeamSetNames=False)

patient.Save()

plan = case.TreatmentPlans[NewPlan_Name]
plan.SetCurrent()


#DELETE POI center CTVp_4000 as not needed anymore
case.PatientModel.PointsOfInterest['center_CTVp_4000'].DeleteRoi()

#set CT name as RCT1
examination = get_current("Examination")

with CompositeAction('Apply image set properties'):

  examination.Name = r"RCT1"

# CompositeAction ends 

#Delete exam group som if exists already?
#case.DeleteExaminationGroup(ExaminationGroupName=r"Simulated organ motion")

#Delete ROI called FLASH if already exists
rois = case.PatientModel.RegionsOfInterest
roi_list = []

for roi in rois:
    roi_list.append(roi.Name)

roi_name1 = 'FLASH'

if (roi_name1 in roi_list):
    case.PatientModel.RegionsOfInterest['FLASH'].DeleteRoi()
else:
    pass


#create SOMs CT depending on right or left plan name
plan = get_current ("Plan")
if plan.Name == 'L_Breast_VMAT': 
    case.GenerateOrganMotionExaminationGroup(SourceExaminationName=r"RCT1", ExaminationGroupName=r"Simulated organ motion", MotionRoiName=r"CTVp_4000", FixedRoiNames=[r"Ipsi_Lung"], OrganUncertaintySettings={ 'Superior': 0, 'Inferior': 0, 'Anterior': 1.5, 'Posterior': 0, 'Right': 0, 'Left': 1.5 }, OnlySimulateMaxOrganMotion=True)
elif plan.Name == 'R_Breast_VMAT':
    case.GenerateOrganMotionExaminationGroup(SourceExaminationName=r"RCT1", ExaminationGroupName=r"Simulated organ motion", MotionRoiName=r"CTVp_4000", FixedRoiNames=[r"Ipsi_Lung"], OrganUncertaintySettings={ 'Superior': 0, 'Inferior': 0, 'Anterior': 1.5, 'Posterior': 0, 'Right': 1.5, 'Left': 0 }, OnlySimulateMaxOrganMotion=True)    

#create FLASH ROI dpending on right or left plan name
if plan.Name == 'L_Breast_VMAT':
    examination = case.Examinations["CTVp_4000 (R-L: 1.50, I-S: 0.00, P-A: 0.00)"]
    examination.SetPrimary()

    with CompositeAction('ROI algebra (SkinSOM1, Image set: CTVp_4000 (R-L: 1.50, I-S: 0.00, P-A: 0.00))'):

        retval_0 = case.PatientModel.CreateRoi(Name=r"SkinSOM1", Color="Yellow", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_0.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_0.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    case.PatientModel.CopyRoiGeometries(SourceExamination=examination, TargetExaminationNames=[r"RCT1"], RoiNames=[r"SkinSOM1"])

    examination = case.Examinations["CTVp_4000 (R-L: 0.00, I-S: 0.00, P-A: 1.50)"]
    examination.SetPrimary()

    with CompositeAction('ROI algebra (SkinSOM2, Image set: CTVp_4000 (R-L: 0.00, I-S: 0.00, P-A: 1.50))'):

        retval_1 = case.PatientModel.CreateRoi(Name=r"SkinSOM2", Color="Green", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_1.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_1.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    case.PatientModel.CopyRoiGeometries(SourceExamination=examination, TargetExaminationNames=[r"RCT1"], RoiNames=[r"SkinSOM2"])

    examination = case.Examinations["CTVp_4000 (R-L: 0.79, I-S: 0.00, P-A: 1.28)"]
    examination.SetPrimary()


    with CompositeAction('ROI algebra (SkinSOM3, Image set: CTVp_4000 (R-L: 0.79, I-S: 0.00, P-A: 1.28))'):

        retval_2 = case.PatientModel.CreateRoi(Name=r"SkinSOM3", Color="Blue", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_2.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_2.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    case.PatientModel.CopyRoiGeometries(SourceExamination=examination, TargetExaminationNames=[r"RCT1"], RoiNames=[r"SkinSOM3"])


    examination = case.Examinations["RCT1"]
    examination.SetPrimary()

    with CompositeAction('ROI algebra (FLASH, Image set: RCT1)'):

        retval_3 = case.PatientModel.CreateRoi(Name=r"FLASH", Color="Magenta", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_3.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"SkinSOM1", r"SkinSOM2", r"SkinSOM3"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Subtraction", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_3.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    with CompositeAction('Delete ROI (SkinSOM1, SkinSOM2, SkinSOM3)'):

        retval_0.DeleteRoi()

        retval_1.DeleteRoi()

        retval_2.DeleteRoi()

        # CompositeAction ends 


    case.PatientModel.StructureSets['RCT1'].SimplifyContours(RoiNames=[r"FLASH"], RemoveHoles3D=False, RemoveSmallContours=True, AreaThreshold=3.0, ReduceMaxNumberOfPointsInContours=False, MaxNumberOfPoints=None, CreateCopyOfRoi=False, ResolveOverlappingContours=False)

elif plan.Name == 'R_Breast_VMAT':
    examination = case.Examinations["CTVp_4000 (R-L: -1.50, I-S: 0.00, P-A: 0.00)"]
    examination.SetPrimary()

    with CompositeAction('ROI algebra (SkinSOM1, Image set: CTVp_4000 (R-L: -1.50, I-S: 0.00, P-A: 0.00))'):

        retval_0 = case.PatientModel.CreateRoi(Name=r"SkinSOM1", Color="Yellow", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_0.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_0.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    case.PatientModel.CopyRoiGeometries(SourceExamination=examination, TargetExaminationNames=[r"RCT1"], RoiNames=[r"SkinSOM1"])

    examination = case.Examinations["CTVp_4000 (R-L: 0.00, I-S: 0.00, P-A: 1.50)"]
    examination.SetPrimary()

    with CompositeAction('ROI algebra (SkinSOM2, Image set: CTVp_4000 (R-L: 0.00, I-S: 0.00, P-A: 1.50))'):

        retval_1 = case.PatientModel.CreateRoi(Name=r"SkinSOM2", Color="Green", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_1.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_1.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    case.PatientModel.CopyRoiGeometries(SourceExamination=examination, TargetExaminationNames=[r"RCT1"], RoiNames=[r"SkinSOM2"])

    examination = case.Examinations["CTVp_4000 (R-L: -0.79, I-S: 0.00, P-A: 1.28)"]
    examination.SetPrimary()


    with CompositeAction('ROI algebra (SkinSOM3, Image set: CTVp_4000 (R-L: -0.79, I-S: 0.00, P-A: 1.28))'):

        retval_2 = case.PatientModel.CreateRoi(Name=r"SkinSOM3", Color="Blue", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_2.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="None", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_2.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    case.PatientModel.CopyRoiGeometries(SourceExamination=examination, TargetExaminationNames=[r"RCT1"], RoiNames=[r"SkinSOM3"])


    examination = case.Examinations["RCT1"]
    examination.SetPrimary()

    with CompositeAction('ROI algebra (FLASH, Image set: RCT1)'):

        retval_3 = case.PatientModel.CreateRoi(Name=r"FLASH", Color="Magenta", Type="Organ", TissueName=None, RbeCellTypeName=None, RoiMaterial=None)

        retval_3.SetAlgebraExpression(ExpressionA={ 'Operation': "Union", 'SourceRoiNames': [r"SkinSOM1", r"SkinSOM2", r"SkinSOM3"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ExpressionB={ 'Operation': "Union", 'SourceRoiNames': [r"Skin"], 'MarginSettings': { 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 } }, ResultOperation="Subtraction", ResultMarginSettings={ 'Type': "Expand", 'Superior': 0, 'Inferior': 0, 'Anterior': 0, 'Posterior': 0, 'Right': 0, 'Left': 0 })

        retval_3.UpdateDerivedGeometry(Examination=examination, Algorithm="Auto")

        # CompositeAction ends 


    with CompositeAction('Delete ROI (SkinSOM1, SkinSOM2, SkinSOM3)'):

        retval_0.DeleteRoi()

        retval_1.DeleteRoi()

        retval_2.DeleteRoi()

        # CompositeAction ends 


    case.PatientModel.StructureSets['RCT1'].SimplifyContours(RoiNames=[r"FLASH"], RemoveHoles3D=False, RemoveSmallContours=True, AreaThreshold=3.0, ReduceMaxNumberOfPointsInContours=False, MaxNumberOfPoints=None, CreateCopyOfRoi=False, ResolveOverlappingContours=False)

#Reset Optimisation andf set settings
plan.PlanOptimizations[0].ResetOptimization()

plan.PlanOptimizations[0].OptimizationParameters.Algorithm.OptimalityTolerance = 1e-7


plan.PlanOptimizations[0].OptimizationParameters.Algorithm.MaxNumberOfIterations = 40


plan.PlanOptimizations[0].OptimizationParameters.DoseCalculation.ComputeFinalDose = True

#addds robustness ticks to relevant functions
constituent_functions = [cf for cf in plan.PlanOptimizations[0].Objective.ConstituentFunctions ]
    
for index,cf in enumerate(constituent_functions):
    
    func_type = ""
    roi_Name = cf.ForRegionOfInterest.Name
    roi_list = ['PTVp_4000','PTVn_LN_IMN','PTVn_LN_Ax_L1','PTVn_LN_Ax_L2','PTVn_LN_Ax_L3','PTVn_LN_Ax_L4','PTVn_LN_Interpect','PTVp_4000-PTVp_4800_DVH','PTVp_4000-PTVp_4800_DVH+1cm']
    func_type1 = ['MinDose', 'MaxDose']
    if hasattr( cf.DoseFunctionParameters, 'FunctionType' ):
        func_type = cf.DoseFunctionParameters.FunctionType
    elif hasattr( cf.DoseFunctionParameters, 'LowDoseDistance' ):
        func_type = "DoseFallOff"   # some special case that had no FunctionType attribute in StateTree.
    else:
        raise PlanException("Unrecognized objective type for objective number {}".format(index))     
    if func_type in func_type1 and roi_Name in roi_list:
        cf.UseRobustness = True
    else:
        cf.UseRobustness = False

#set robustness dataset for left
if plan.Name == 'L_Breast_VMAT': 
    plan.PlanOptimizations[0].OptimizationParameters.SaveRobustnessParameters(PositionUncertaintyAnterior=0, PositionUncertaintyPosterior=0, PositionUncertaintySuperior=0, PositionUncertaintyInferior=0, PositionUncertaintyLeft=0, PositionUncertaintyRight=0, DensityUncertainty=0, PositionUncertaintySetting="Universal", IndependentLeftRight=True, IndependentAnteriorPosterior=True, IndependentSuperiorInferior=True, ComputeExactScenarioDoses=False, NamesOfNonPlanningExaminations=[r"CTVp_4000 (R-L: 1.50, I-S: 0.00, P-A: 0.00)", r"CTVp_4000 (R-L: 0.00, I-S: 0.00, P-A: 1.50)", r"CTVp_4000 (R-L: 0.79, I-S: 0.00, P-A: 1.28)"])
elif plan.Name == 'R_Breast_VMAT':
    plan.PlanOptimizations[0].OptimizationParameters.SaveRobustnessParameters(PositionUncertaintyAnterior=0, PositionUncertaintyPosterior=0, PositionUncertaintySuperior=0, PositionUncertaintyInferior=0, PositionUncertaintyLeft=0, PositionUncertaintyRight=0, DensityUncertainty=0, PositionUncertaintySetting="Universal", IndependentLeftRight=True, IndependentAnteriorPosterior=True, IndependentSuperiorInferior=True, ComputeExactScenarioDoses=False, NamesOfNonPlanningExaminations=[r"CTVp_4000 (R-L: -1.50, I-S: 0.00, P-A: 0.00)", r"CTVp_4000 (R-L: 0.00, I-S: 0.00, P-A: 1.50)", r"CTVp_4000 (R-L: -0.79, I-S: 0.00, P-A: 1.28)"])

#increase dose grid size ant to account for potential target on defomred CT outside original grid. this happened to me once before and would not optimise until i expanded grid. this is when the most ant part of skin roi is the breast and not chin.
dose_grid = plan.GetDoseGrid()

plan.UpdateDoseGrid(Corner={ 'x': dose_grid.Corner.x, 'y': dose_grid.Corner.y - 1.5, 'z': dose_grid.Corner.z }, VoxelSize={ 'x': dose_grid.VoxelSize.x, 'y': dose_grid.VoxelSize.y, 'z': dose_grid.VoxelSize.z }, NumberOfVoxels={ 'x': dose_grid.NrVoxels.x, 'y': dose_grid.NrVoxels.y + 10, 'z': dose_grid.NrVoxels.z })

plan.TreatmentCourse.TotalDose.UpdateDoseGridStructures()

#autoscale off, optimise cold start, warm start, autoscale on, turn flash off and save patient
plan.PlanOptimizations[0].AutoScaleToPrescription = False
  
plan.PlanOptimizations[0].RunOptimization()

plan.PlanOptimizations[0].RunOptimization()

patient.Save()

plan.PlanOptimizations[0].AutoScaleToPrescription = True

patient.SetRoiVisibility(RoiName='FLASH', IsVisible=False)

#patient.Save()

#renames plan and beamset to include SIB if PX IS 48Gy
beam_set = get_current("BeamSet")
px_dose = beam_set.Prescription.PrimaryDosePrescription.DoseValue

if px_dose == 4800:
    if plan.Name == 'L_Breast_VMAT': 
        plan.Name = 'LBreast_VMATSIB'
if px_dose == 4800:
    if plan.Name == 'R_Breast_VMAT': 
        plan.Name = 'RBreast_VMATSIB'
else:
    pass
plan.BeamSets[0].DicomPlanLabel = plan.Name

patient.Save()

rtn = ctypes.windll.user32.MessageBoxA(0, 'Script Complete. Remember once plan complete to make ROI isodoses 32Gy, 36Gy, 38Gy and untick robustness before exporting', 'Breast VMAT ROBUST v' + str(script_version), 0)
