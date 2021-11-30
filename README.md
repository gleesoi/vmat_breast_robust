# vmat_breast_robust
RS V10A planning scripts - robust opt for vmat breast and nodes
BREAST VMAT ROBUST v1.0
CPython
Ian Gleeson
June 2021



This script will complete all the robust parts of VMAT planning and create the FLASH ROI. It runs for both left or right sided and will label the plan correctly even if planner has labelled the laterality incorrectly.
 
The script will complete steps 7.4.6 to 7.4.18 (begins with copying the current plan) and also make FLASH ROI (7.4.22) in tp3.01. It takes about 11minutes for script to complete (as it includes a cold start and a warm start optimisation).
 
Requirements will be that the plan name must be called either R_Breast_Initial or L_Breast_Initial and no plan called L_Breast_VMAT or R_Breast_VMAT in case (as this will be the copied plan name) as per tp3.01. It should flag anything else that it’s not happy with for example such not having the ROIs ‘Skin’, ‘Ipsi_Lung’, ‘CTVp_4000’ or already having an organ motion simulated exam group etc.
 
I have made it so it works now with newly approved breast ROI names – e.g. PTVn_LN_Ax_L1
 
I have tested it on right and left case and each part of code and am sure it will also run on a VMAT nodal SIB
