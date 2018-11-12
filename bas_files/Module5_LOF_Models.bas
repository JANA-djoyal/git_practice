Attribute VB_Name = "Module5_LOF_Models"
'Ensure text comparisons are done ignoring case
Option Compare Text
'Enable Option Explicit
Option Explicit

Public Function LOF_Model(Inspection_Interval As Double, i As Long, _
    access_dict As Scripting.Dictionary, locate_dict As Scripting.Dictionary, _
    actuate_dict As Scripting.Dictionary, leak_dict As Scripting.Dictionary, _
    loc_code_dict As Scripting.Dictionary, press_dict As Scripting.Dictionary, _
    size_dict As Scripting.Dictionary, type_dict As Scripting.Dictionary) As Variant

    'General variables for loops etc.
    Dim j As Long
    'Initialize values of general variables
    j = 0
    'Initialize input variables
    Dim valvedata As Variant
    
    'Read in line of valve data
     valvedata = Worksheets("Valve_List").Range("A" & (i + 5), "AD" & (i + 5)).Value

' ------------------------------------------------------

    ' Initializing factors for proportional hazards model
    Dim cur_use As String
    Dim cur_loc_class As String
    Dim cur_loc_code As String
    Dim cur_Size As String
    Dim cur_type As String
    Dim cur_Mat As String
    Dim cur_Press As String
    Dim cur_position As String
    Dim cur_accesshist As Double
    Dim cur_locatehist As Double
    Dim cur_actuatehist As Double
    Dim cur_leakhist As Double
    Dim cur_age As Double
    Dim cur_install_year As Double
    
    
    'Initialize intermediate variables for use in calculations
    Dim cur_access_PHM_Factor As Double
    Dim cur_locate_PHM_Factor As Double
    Dim cur_actuate_PHM_Factor As Double
    Dim leak_PHM_Factor As Double
    Dim cur_access_base_rate As Double
    Dim cur_locate_base_rate As Double
    Dim actuate_base_rate_immediate As Double
    Dim actuate_base_rate_maxyr As Double
    Dim leak_base_rate_immediate As Double
    Dim leak_base_rate_maxyr As Double
    Dim cur_R_Access As Double
    Dim cur_R_Locate As Double
    Dim avg_R_Actuate As Double
    Dim avg_R_Leak As Double
    Dim LoF_Operate As Double
    Dim LoF_Leak As Double
    
    'Model Applications
        'read values for PHM
        If valvedata(1, 10) = "" Then
            If valvedata(1, 10) = "STEEL" Or valvedata(1, 10) = "CASTIRON" Then
                cur_install_year = 1975
            Else
                cur_install_year = 1995
            End If
        Else
            cur_install_year = valvedata(1, 10)
        End If
        
        'the following may throw error if blank - check/set default
        cur_loc_class = valvedata(1, 11) 'this will have to reference a dictionary like the ones below I think once we get values
        
        If valvedata(1, 21) = "" Then
            cur_loc_code = "Blvd/Ease/Prk/Schl/Other"
        Else
            cur_loc_code = loc_code_dict(valvedata(1, 21)) 'figure out how to deal with this given the new criticality field
        End If
        
        If valvedata(1, 9) = "" Then
            cur_loc_code = "<4in"
        Else
            cur_loc_code = loc_code_dict(valvedata(1, 9))
        End If
        
        cur_type = type_dict(valvedata(1, 18)) 'can probably use raw output just change dict/table in "Modelling Categories" sheet
        cur_Press = press_dict(valvedata(1, 15)) 'check all captured - what to do with blanks
        cur_position = valvedata(1, 5)
        cur_accesshist = valvedata(1, 27)
        cur_locatehist = valvedata(1, 28)
        cur_actuatehist = valvedata(1, 29)
        cur_leakhist = valvedata(1, 30)
               
        'Access Model Calculations
 
        cur_access_base_rate = (Inspection_Interval / access_dict("Eta")) ^ access_dict("Beta")
        
        cur_access_PHM_Factor = Exp(access_dict(cur_loc_class) _
                                     + access_dict(cur_loc_code) _
                                     + access_dict(cur_Size) _
                                     + access_dict(cur_Press) _
                                     + access_dict(cur_type) _
                                     + cur_accesshist * access_dict("Index"))
  
        cur_R_Access = Exp(-cur_access_base_rate * cur_access_PHM_Factor)
            
        'Locate Model Calculations
        
        cur_locate_base_rate = (Inspection_Interval / locate_dict("Eta")) ^ locate_dict("Beta")
        
        cur_locate_PHM_Factor = Exp(locate_dict(cur_loc_class) _
                                     + locate_dict(cur_loc_code) _
                                     + locate_dict(cur_Size) _
                                     + locate_dict(cur_Press) _
                                     + locate_dict(cur_type) _
                                     + cur_locatehist * locate_dict("Index"))
        
        cur_R_Locate = Exp(-cur_locate_base_rate * cur_locate_PHM_Factor)
        
        
        'Actuation Model Calculations
        cur_age = Year(Date) - cur_install_year
        
        actuate_base_rate_immediate = ((Inspection_Interval + cur_age) / actuate_dict("Eta")) ^ actuate_dict("Beta") - (cur_age / actuate_dict("Eta")) ^ actuate_dict("Beta")
        actuate_base_rate_maxyr = ((cur_age + 10) / actuate_dict("Eta")) ^ actuate_dict("Beta") _
                                - ((cur_age + 10 - Inspection_Interval) / actuate_dict("Eta")) ^ actuate_dict("Beta")
        cur_actuate_PHM_Factor = Exp(actuate_dict(cur_loc_class) _
                                     + actuate_dict(cur_loc_code) _
                                     + actuate_dict(cur_Size) _
                                     + actuate_dict(cur_Press) _
                                     + actuate_dict(cur_type) _
                                     + cur_actuatehist * actuate_dict("Index"))
       
        avg_R_Actuate = (Exp(-actuate_base_rate_immediate * cur_actuate_PHM_Factor) + Exp(-actuate_base_rate_maxyr * cur_actuate_PHM_Factor)) / 2
    
        'Leak Model Calculations
        
        leak_base_rate_immediate = ((Inspection_Interval + cur_age) / leak_dict("Eta")) ^ leak_dict("Beta") - (cur_age / leak_dict("Eta")) ^ leak_dict("Beta")
        leak_base_rate_maxyr = ((cur_age + 10) / leak_dict("Eta")) ^ leak_dict("Beta") _
                                - ((cur_age + 10 - Inspection_Interval) / leak_dict("Eta")) ^ leak_dict("Beta")
        
        leak_PHM_Factor = Exp(leak_dict(cur_loc_class) _
                                     + leak_dict(cur_loc_code) _
                                     + leak_dict(cur_Size) _
                                     + leak_dict(cur_Press) _
                                     + leak_dict(cur_type) _
                                     + cur_leakhist * leak_dict("Index"))
        
        avg_R_Leak = (Exp(-leak_base_rate_immediate * leak_PHM_Factor) + Exp(-leak_base_rate_maxyr * leak_PHM_Factor)) / 2
        
        
        ' Overall Lof For Operational Failures Over Inspection Interval
        LoF_Operate = 1 - (cur_R_Access * cur_R_Locate * avg_R_Actuate)
    
        ' Overall Lof For Leaks Over Inspection Interval
        LoF_Leak = 1 - avg_R_Leak
        
        'output values
        Dim Outputs(2) As Double
        Outputs(1) = LoF_Operate
        Outputs(2) = LoF_Leak
        LOF_Model = Outputs


End Function

