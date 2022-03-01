from connect import *
patient = get_current("Patient")
plan = get_current('Plan')
case = get_current('Case')

# Access the plan dose
plan_dose = plan.TreatmentCourse.TotalDose

# Define the threshold level
threshold_level = 3200

# Create a new ROI and create its geometry from the plan dose
# and the threshold level
# Define the name of the ROI
roi_name = '32Gy isodose '
roi = case.PatientModel.CreateRoi(Name = roi_name, Color = 'Blue', Type = 'Control')
roi.CreateRoiGeometryFromDose(DoseDistribution = plan_dose, ThresholdLevel = threshold_level)

# Define the threshold level
threshold_level2 = 3600

# Create a new ROI and create its geometry from the plan dose
# and the threshold level
# Define the name of the ROI
roi_name = '36Gy isodose '
roi = case.PatientModel.CreateRoi(Name = roi_name, Color = 'Green', Type = 'Control')
roi.CreateRoiGeometryFromDose(DoseDistribution = plan_dose, ThresholdLevel = threshold_level2)

# Define the threshold level
threshold_level3 = 3800

# Create a new ROI and create its geometry from the plan dose
# and the threshold level
# Define the name of the ROI
roi_name = '38Gy isodose '
roi = case.PatientModel.CreateRoi(Name = roi_name, Color = 'Brown', Type = 'Control')
roi.CreateRoiGeometryFromDose(DoseDistribution = plan_dose, ThresholdLevel = threshold_level3)

##turn off robust settings so can export ok
plan.PlanOptimizations[0].OptimizationParameters.SaveRobustnessParameters(PositionUncertaintyAnterior=0, PositionUncertaintyPosterior=0, PositionUncertaintySuperior=0, PositionUncertaintyInferior=0, PositionUncertaintyLeft=0, PositionUncertaintyRight=0, DensityUncertainty=0, PositionUncertaintySetting="Universal", IndependentLeftRight=True, IndependentAnteriorPosterior=True, IndependentSuperiorInferior=True, ComputeExactScenarioDoses=False, NamesOfNonPlanningExaminations=[])

##refresh contours grid

plan.TreatmentCourse.TotalDose.UpdateDoseGridStructures()

##save
patient.Save()