Attribute VB_Name = "Module1_Optimization"
'Ensure text comparisons are done ignoring case
Option Compare Text
'Enable Option Explicit
Option Explicit

Sub JVIO_Optimization()

'-----------------SETUP-------------------------

    'Allows VBA to edit sheet while keeping interface protected
    Worksheets("Inputs").Unprotect "QS"
    Worksheets("Inputs").Protect "QS", userinterfaceonly:=True
    Worksheets("Results_List").Unprotect "QS"
    Worksheets("Results_List").Protect "QS", userinterfaceonly:=True, AllowFiltering:=True
    Worksheets("Levelled_Inspections").Unprotect "QS"
    Worksheets("Levelled_Inspections").Protect "QS", userinterfaceonly:=True, AllowFiltering:=True
    
    'Remove all buttons while running to avoid issues (print still works)
    Worksheets("Inputs").Shapes("CommandButton1").Visible = False
    Worksheets("Inputs").Shapes("CommandButton2").Visible = False
    Worksheets("Inputs").Shapes("CommandButton3").Visible = False
        
    'Initiate Statusbar
    application.DisplayStatusBar = True
        
    'Lock User Inputs
    Worksheets("Inputs").Range("B7:C8").Locked = True
    Worksheets("Inputs").Range("B9:F10").Locked = True
    Worksheets("Inputs").Range("B14:C17").Locked = True
    Worksheets("Inputs").Range("C23:C24").Locked = True
    Worksheets("Inputs").Range("H15:H25").Locked = True
    Worksheets("Inputs").Range("C30:D31").Locked = True
    Worksheets("Inputs").Range("E30:F31").Locked = True
    Worksheets("Inputs").Range("C36:L46").Locked = True
    Worksheets("Inputs").Range("C51:D52").Locked = True
    Worksheets("Inputs").Range("C57:F67").Locked = True

    'Clear Output Sheets
    With Worksheets("Results_List")
    .Rows("7:" & .Rows.Count).ClearContents
    End With
    
    With Worksheets("Levelled_Inspections")
    .Rows("7:" & .Rows.Count).ClearContents
    End With
    
    'Get Workbook Name - avoids issues of leaving current book
    Dim curbook
    curbook = ActiveWorkbook.Name
    
    'Error Handling
    On Error GoTo ErrHandler
    application.EnableCancelKey = xlErrorHandler


    
 ' -----------------INITIALIZING VARIABLES AND VALUES-------------------------------------
    
    'General variables for loops etc.
    Dim i As Long
    Dim j As Long
    'Initialize values of general variables
    i = 0
    j = 0
    'Initialize input variables
    Dim numrecords As Long
    Dim valves As Variant
    
    'Read in all valve data
    valves = Worksheets("Valve_List").Range("A6:AD20000").Value
    
    'Read in control parameters (survey intervals and valve criticality factors)
    Dim min_survey_int
    Set min_survey_int = CreateObject("Scripting.Dictionary")
    With min_survey_int
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("H15").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("H16").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("H17").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("H18").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("H19").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("H20").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("H21").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("H22").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("H23").Value
    End With
    
    Dim funct_incident_factors_iso
    Set funct_incident_factors_iso = CreateObject("Scripting.Dictionary")
    With funct_incident_factors_iso
        .Add "Enhanced", Worksheets("Inputs").Range("C36").Value
        .Add "Special", Worksheets("Inputs").Range("C37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("C38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("C39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("C40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("C41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("C42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("C43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("C44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("C45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("C46").Value
    End With
    
    Dim funct_incident_factors_ha
    Set funct_incident_factors_ha = CreateObject("Scripting.Dictionary")
    With funct_incident_factors_ha
        .Add "Enhanced", Worksheets("Inputs").Range("D36").Value
        .Add "Special", Worksheets("Inputs").Range("D37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("D38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("D39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("D40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("D41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("D42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("D43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("D44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("D45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("D46").Value
    End With
    
    Dim funct_consequence_factors_Iso
    Set funct_consequence_factors_Iso = CreateObject("Scripting.Dictionary")
    With funct_consequence_factors_Iso
        .Add "Enhanced", Worksheets("Inputs").Range("E36").Value
        .Add "Special", Worksheets("Inputs").Range("E37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("E38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("E39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("E40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("E41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("E42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("E43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("E44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("E45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("E46").Value
    End With
    
    Dim funct_consequence_factors_ha
    Set funct_consequence_factors_ha = CreateObject("Scripting.Dictionary")
    With funct_consequence_factors_ha
        .Add "Enhanced", Worksheets("Inputs").Range("F36").Value
        .Add "Special", Worksheets("Inputs").Range("F37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("F38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("F39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("F40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("F41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("F42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("F43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("F44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("F45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("F46").Value
    End With
    
    Dim funct_failure_factors_Iso
    Set funct_failure_factors_Iso = CreateObject("Scripting.Dictionary")
    With funct_failure_factors_Iso
        .Add "Enhanced", Worksheets("Inputs").Range("G36").Value
        .Add "Special", Worksheets("Inputs").Range("G37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("G38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("G39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("G40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("G41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("G42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("G43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("G44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("G45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("G46").Value
    End With
    
    Dim funct_failure_factors_ha
    Set funct_failure_factors_ha = CreateObject("Scripting.Dictionary")
    With funct_failure_factors_ha
        .Add "Enhanced", Worksheets("Inputs").Range("H36").Value
        .Add "Special", Worksheets("Inputs").Range("H37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("H38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("H39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("H40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("H41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("H42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("H43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("H44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("H45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("H46").Value
    End With
    
'    Dim funct_opening_factors_Iso
'    Set funct_opening_factors_Iso = CreateObject("Scripting.Dictionary")
'    With funct_opening_factors_Iso
'        .Add "Enhanced", Worksheets("Inputs").Range("K36").Value
'        .Add "Special", Worksheets("Inputs").Range("K37").Value
'        .Add "Urban Crit - AG", Worksheets("Inputs").Range("K38").Value
'        .Add "Urban Crit - BG", Worksheets("Inputs").Range("K39").Value
'        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("K40").Value
'        .Add "Urban - AG", Worksheets("Inputs").Range("K41").Value
'        .Add "Urban - BG", Worksheets("Inputs").Range("K42").Value
'        .Add "Urban - Vault", Worksheets("Inputs").Range("K43").Value
'        .Add "Rural - AG", Worksheets("Inputs").Range("K44").Value
'        .Add "Rural - BG", Worksheets("Inputs").Range("K45").Value
'        .Add "Rural - Vault", Worksheets("Inputs").Range("K46").Value
'    End With
'
'    Dim funct_opening_factors_ha
'    Set funct_opening_factors_ha = CreateObject("Scripting.Dictionary")
'    With funct_opening_factors_ha
'        .Add "Enhanced", Worksheets("Inputs").Range("L36").Value
'        .Add "Special", Worksheets("Inputs").Range("L37").Value
'        .Add "Urban Crit - AG", Worksheets("Inputs").Range("L38").Value
'        .Add "Urban Crit - BG", Worksheets("Inputs").Range("L39").Value
'        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("L40").Value
'        .Add "Urban - AG", Worksheets("Inputs").Range("L41").Value
'        .Add "Urban - BG", Worksheets("Inputs").Range("L42").Value
'        .Add "Urban - Vault", Worksheets("Inputs").Range("L43").Value
'        .Add "Rural - AG", Worksheets("Inputs").Range("L44").Value
'        .Add "Rural - BG", Worksheets("Inputs").Range("L45").Value
'        .Add "Rural - Vault", Worksheets("Inputs").Range("L46").Value
'    End With
    
    Dim funct_opening_consequence_Iso
    Set funct_opening_consequence_Iso = CreateObject("Scripting.Dictionary")
    With funct_opening_consequence_Iso
        .Add "Enhanced", Worksheets("Inputs").Range("I36").Value
        .Add "Special", Worksheets("Inputs").Range("I37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("I38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("I39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("I40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("I41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("I42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("I43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("I44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("I45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("I46").Value
    End With
    
    Dim funct_opening_consequence_ha
    Set funct_opening_consequence_ha = CreateObject("Scripting.Dictionary")
    With funct_opening_consequence_ha
        .Add "Enhanced", Worksheets("Inputs").Range("J36").Value
        .Add "Special", Worksheets("Inputs").Range("J37").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("J38").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("J39").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("J40").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("J41").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("J42").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("J43").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("J44").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("J45").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("J46").Value
    End With
    
    Dim funct_scaling_factors_Iso
    Set funct_scaling_factors_Iso = CreateObject("Scripting.Dictionary")
    With funct_scaling_factors_Iso
        .Add "Enhanced", Worksheets("Defaults").Range("K36").Value
        .Add "Special", Worksheets("Defaults").Range("K37").Value
        .Add "Urban Crit - AG", Worksheets("Defaults").Range("K38").Value
        .Add "Urban Crit - BG", Worksheets("Defaults").Range("K39").Value
        .Add "Urban Crit - Vault", Worksheets("Defaults").Range("K40").Value
        .Add "Urban - AG", Worksheets("Defaults").Range("K41").Value
        .Add "Urban - BG", Worksheets("Defaults").Range("K42").Value
        .Add "Urban - Vault", Worksheets("Defaults").Range("K43").Value
        .Add "Rural - AG", Worksheets("Defaults").Range("K44").Value
        .Add "Rural - BG", Worksheets("Defaults").Range("K45").Value
        .Add "Rural - Vault", Worksheets("Defaults").Range("K46").Value
    End With
    
    Dim funct_scaling_factors_ha
    Set funct_scaling_factors_ha = CreateObject("Scripting.Dictionary")
    With funct_scaling_factors_ha
        .Add "Enhanced", Worksheets("Defaults").Range("L36").Value
        .Add "Special", Worksheets("Defaults").Range("L37").Value
        .Add "Urban Crit - AG", Worksheets("Defaults").Range("L38").Value
        .Add "Urban Crit - BG", Worksheets("Defaults").Range("L39").Value
        .Add "Urban Crit - Vault", Worksheets("Defaults").Range("L40").Value
        .Add "Urban - AG", Worksheets("Defaults").Range("L41").Value
        .Add "Urban - BG", Worksheets("Defaults").Range("L42").Value
        .Add "Urban - Vault", Worksheets("Defaults").Range("L43").Value
        .Add "Rural - AG", Worksheets("Defaults").Range("L44").Value
        .Add "Rural - BG", Worksheets("Defaults").Range("L45").Value
        .Add "Rural - Vault", Worksheets("Defaults").Range("N46").Value
    End With
    
    Dim leak_incident_factors_iso
    Set leak_incident_factors_iso = CreateObject("Scripting.Dictionary")
    With leak_incident_factors_iso
        .Add "Enhanced", Worksheets("Inputs").Range("C57").Value
        .Add "Special", Worksheets("Inputs").Range("C58").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("C59").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("C60").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("C61").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("C62").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("C63").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("C64").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("C65").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("C66").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("C67").Value
    End With
    
    Dim leak_incident_factors_ha
    Set leak_incident_factors_ha = CreateObject("Scripting.Dictionary")
    With leak_incident_factors_ha
        .Add "Enhanced", Worksheets("Inputs").Range("D57").Value
        .Add "Special", Worksheets("Inputs").Range("D58").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("D59").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("D60").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("D61").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("D62").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("D63").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("D64").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("D65").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("D66").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("D67").Value
    End With
    
    Dim leak_consequence_factors_iso
    Set leak_consequence_factors_iso = CreateObject("Scripting.Dictionary")
    With leak_consequence_factors_iso
        .Add "Enhanced", Worksheets("Inputs").Range("E57").Value
        .Add "Special", Worksheets("Inputs").Range("E58").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("E59").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("E60").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("E61").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("E62").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("E63").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("E64").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("E65").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("E66").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("E67").Value
    End With
    
    Dim leak_consequence_factors_ha
    Set leak_consequence_factors_ha = CreateObject("Scripting.Dictionary")
    With leak_consequence_factors_ha
        .Add "Enhanced", Worksheets("Inputs").Range("F57").Value
        .Add "Special", Worksheets("Inputs").Range("F58").Value
        .Add "Urban Crit - AG", Worksheets("Inputs").Range("F59").Value
        .Add "Urban Crit - BG", Worksheets("Inputs").Range("F60").Value
        .Add "Urban Crit - Vault", Worksheets("Inputs").Range("F61").Value
        .Add "Urban - AG", Worksheets("Inputs").Range("F62").Value
        .Add "Urban - BG", Worksheets("Inputs").Range("F63").Value
        .Add "Urban - Vault", Worksheets("Inputs").Range("F64").Value
        .Add "Rural - AG", Worksheets("Inputs").Range("F65").Value
        .Add "Rural - BG", Worksheets("Inputs").Range("F66").Value
        .Add "Rural - Vault", Worksheets("Inputs").Range("F67").Value
    End With
    
    Dim leak_scaling_factors_iso
    Set leak_scaling_factors_iso = CreateObject("Scripting.Dictionary")
    With leak_scaling_factors_iso
        .Add "Enhanced", Worksheets("Defaults").Range("G57").Value
        .Add "Special", Worksheets("Defaults").Range("G58").Value
        .Add "Urban Crit - AG", Worksheets("Defaults").Range("G59").Value
        .Add "Urban Crit - BG", Worksheets("Defaults").Range("G60").Value
        .Add "Urban Crit - Vault", Worksheets("Defaults").Range("G61").Value
        .Add "Urban - AG", Worksheets("Defaults").Range("G62").Value
        .Add "Urban - BG", Worksheets("Defaults").Range("G63").Value
        .Add "Urban - Vault", Worksheets("Defaults").Range("G64").Value
        .Add "Rural - AG", Worksheets("Defaults").Range("G65").Value
        .Add "Rural - BG", Worksheets("Defaults").Range("G66").Value
        .Add "Rural - Vault", Worksheets("Defaults").Range("G67").Value
    End With
    
    Dim leak_scaling_factors_ha
    Set leak_scaling_factors_ha = CreateObject("Scripting.Dictionary")
    With leak_scaling_factors_ha
        .Add "Enhanced", Worksheets("Defaults").Range("H57").Value
        .Add "Special", Worksheets("Defaults").Range("H58").Value
        .Add "Urban Crit - AG", Worksheets("Defaults").Range("H59").Value
        .Add "Urban Crit - BG", Worksheets("Defaults").Range("H60").Value
        .Add "Urban Crit - Vault", Worksheets("Defaults").Range("H61").Value
        .Add "Urban - AG", Worksheets("Defaults").Range("H62").Value
        .Add "Urban - BG", Worksheets("Defaults").Range("H63").Value
        .Add "Urban - Vault", Worksheets("Defaults").Range("H64").Value
        .Add "Rural - AG", Worksheets("Defaults").Range("H65").Value
        .Add "Rural - BG", Worksheets("Defaults").Range("H66").Value
        .Add "Rural - Vault", Worksheets("Defaults").Range("H67").Value
    End With
    
    'Read in inspection costs
    Dim inspection_cost_BG As Double
    inspection_cost_BG = Worksheets("Inputs").Range("B14").Value
    Dim inspection_cost_AG As Double
    inspection_cost_AG = Worksheets("Inputs").Range("B15").Value
    Dim inspection_cost_Vault As Double
    inspection_cost_Vault = Worksheets("Inputs").Range("B16").Value

' ------------------------------------------------------

    'Read in rate and cost of leak incidents
    Dim iso_leak_incident_rate As Double
    iso_leak_incident_rate = Worksheets("Inputs").Range("C51").Value
    
    Dim ha_leak_incident_rate As Double
    ha_leak_incident_rate = Worksheets("Inputs").Range("D51").Value
    
    Dim iso_leak_cost As Double
    iso_leak_cost = Worksheets("Inputs").Range("C52").Value
    
    Dim ha_leak_cost As Double
    ha_leak_cost = Worksheets("Inputs").Range("D52").Value
    
    'Read in rate and cost of functional failure incidents
    
    Dim ISO_ops_incident_rate As Double
    ISO_ops_incident_rate = Worksheets("Inputs").Range("C30").Value
    
    Dim HA_ops_incident_rate As Double
    HA_ops_incident_rate = Worksheets("Inputs").Range("D30").Value
    
    Dim ISO_open_incident_rate As Double
    ISO_open_incident_rate = Worksheets("Inputs").Range("E30").Value
    
    Dim HA_open_incident_rate As Double
    HA_open_incident_rate = Worksheets("Inputs").Range("F30").Value
    
    Dim ISO_ops_fail_cost As Double
    ISO_ops_fail_cost = Worksheets("Inputs").Range("C31").Value
    
    Dim HA_ops_fail_cost As Double
    HA_ops_fail_cost = Worksheets("Inputs").Range("D31").Value
    
    Dim ISO_open_fail_cost As Double
    ISO_open_fail_cost = Worksheets("Inputs").Range("E31").Value
    
    Dim HA_open_fail_cost As Double
    HA_open_fail_cost = Worksheets("Inputs").Range("F31").Value
    
    ' Read in Winter info
    Dim Do_Winter As String
    Do_Winter = Worksheets("Inputs").Range("C23").Value

    Dim Winter_Cutoff As Double
    Winter_Cutoff = Worksheets("Inputs").Range("C24").Value
' ------------------------------------------------------
    
    'read in Above Ground(AG) access model parameters
    Dim AG_Access_Model_Params As New Scripting.Dictionary
    With AG_Access_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("C3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("C4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("C5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("C6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("C7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("C8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("C9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("C10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("C11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("C12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("C13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("C14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("C17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("C18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("C19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("C20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("C21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("C22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("C23").Value
    End With
    
    'read in AG locate model parameters
    Dim AG_Locate_Model_Params As New Scripting.Dictionary
    With AG_Locate_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("E3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("E4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("E5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("E6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("E7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("E8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("E9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("E10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("E11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("E12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("E13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("E14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("E17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("E18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("E19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("E20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("E21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("E22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("E23").Value
    End With
    
    'read in AG actuate model parameters
    Dim AG_Actuate_Model_Params As New Scripting.Dictionary
    With AG_Actuate_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("G3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("G4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("G5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("G6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("G7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("G8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("G9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("G10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("G11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("G12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("G13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("G14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("G17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("G18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("G19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("G20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("G21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("G22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("G23").Value
    End With
    
    'read in AG leak model parameters
    Dim AG_Leak_Model_Params As New Scripting.Dictionary
    With AG_Leak_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("I3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("I4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("I5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("I6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("I7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("I8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("I9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("I10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("I11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("I12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("I13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("I14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("I17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("I18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("I19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("I20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("I21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("I22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("I23").Value
    End With
    
    'read in Below Ground/Vault (BG) access model parameters
    Dim BG_Access_Model_Params As New Scripting.Dictionary
    With BG_Access_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("N3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("N4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("N5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("N6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("N7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("N8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("N9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("N10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("N11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("N12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("N13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("N14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("N17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("N18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("N19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("N20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("N21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("N22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("N23").Value
    End With
    
    'read in BG Locate model parameters
    Dim BG_Locate_Model_Params As New Scripting.Dictionary
    With BG_Locate_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("P3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("P4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("P5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("P6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("P7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("P8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("P9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("P10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("P11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("P12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("P13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("P14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("P17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("P18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("P19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("P20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("P21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("P22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("P23").Value
    End With
    
    'read in BG Actuate model parameters
    Dim BG_Actuate_Model_Params As New Scripting.Dictionary
    With BG_Actuate_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("R3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("R4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("R5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("R6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("R7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("R8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("R9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("R10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("R11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("R12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("R13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("R14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("R17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("R18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("R19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("R20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("R21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("R22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("R23").Value
    End With
    
    'read in BG Leak model parameters
    Dim BG_Leak_Model_Params As New Scripting.Dictionary
    With BG_Leak_Model_Params
        .Add "Beta", Worksheets("Recommended_Parameters").Range("T3").Value
        .Add "Eta", Worksheets("Recommended_Parameters").Range("T4").Value
        .Add "Frequency", Worksheets("Recommended_Parameters").Range("T5").Value
        .Add "Index", Worksheets("Recommended_Parameters").Range("T6").Value
        .Add "RURAL", Worksheets("Recommended_Parameters").Range("T7").Value
        .Add "URBAN", Worksheets("Recommended_Parameters").Range("T8").Value
        .Add "URBAN-CRIT", Worksheets("Recommended_Parameters").Range("T9").Value
        .Add "Station", Worksheets("Recommended_Parameters").Range("T10").Value
        .Add "Blvd/Ease/Prk/Schl/Other", Worksheets("Recommended_Parameters").Range("T11").Value
        .Add "Rd/Int/Side/Ln/Ally/Update", Worksheets("Recommended_Parameters").Range("T12").Value
        .Add "<4in", Worksheets("Recommended_Parameters").Range("T13").Value
        .Add ">4in", Worksheets("Recommended_Parameters").Range("T14").Value
        .Add "MP", Worksheets("Recommended_Parameters").Range("T17").Value
        .Add "IP", Worksheets("Recommended_Parameters").Range("T18").Value
        .Add "Ball Valve", Worksheets("Recommended_Parameters").Range("T19").Value
        .Add "Butterfly Valve", Worksheets("Recommended_Parameters").Range("T20").Value
        .Add "Gate Valve", Worksheets("Recommended_Parameters").Range("T21").Value
        .Add "Plug Valve", Worksheets("Recommended_Parameters").Range("T22").Value
        .Add "Other", Worksheets("Recommended_Parameters").Range("T23").Value
    End With

'-------------------------------------------------------
'Read in categorizations (i.e. classifying sizes into 2 groups, classifying locationsin to 3 etc.)
    Dim Loc_Code_Cat As New Scripting.Dictionary
    With Loc_Code_Cat
        .Add Worksheets("Modelling Categories").Range("A2").Value, Worksheets("Modelling Categories").Range("B2").Value
        .Add Worksheets("Modelling Categories").Range("A3").Value, Worksheets("Modelling Categories").Range("B3").Value
        .Add Worksheets("Modelling Categories").Range("A4").Value, Worksheets("Modelling Categories").Range("B4").Value
        .Add Worksheets("Modelling Categories").Range("A5").Value, Worksheets("Modelling Categories").Range("B5").Value
        .Add Worksheets("Modelling Categories").Range("A6").Value, Worksheets("Modelling Categories").Range("B6").Value
        .Add Worksheets("Modelling Categories").Range("A7").Value, Worksheets("Modelling Categories").Range("B7").Value
        .Add Worksheets("Modelling Categories").Range("A8").Value, Worksheets("Modelling Categories").Range("B8").Value
        .Add Worksheets("Modelling Categories").Range("A9").Value, Worksheets("Modelling Categories").Range("B9").Value
        .Add Worksheets("Modelling Categories").Range("A10").Value, Worksheets("Modelling Categories").Range("B10").Value
        .Add Worksheets("Modelling Categories").Range("A11").Value, Worksheets("Modelling Categories").Range("B11").Value
    End With
    
    Dim Pressure_Cat As New Scripting.Dictionary
    With Pressure_Cat
        .Add Worksheets("Modelling Categories").Range("D2").Value, Worksheets("Modelling Categories").Range("E2").Value
        .Add Worksheets("Modelling Categories").Range("D3").Value, Worksheets("Modelling Categories").Range("E3").Value
        .Add Worksheets("Modelling Categories").Range("D4").Value, Worksheets("Modelling Categories").Range("E4").Value
        .Add Worksheets("Modelling Categories").Range("D5").Value, Worksheets("Modelling Categories").Range("E5").Value
        .Add Worksheets("Modelling Categories").Range("D6").Value, Worksheets("Modelling Categories").Range("E6").Value
        .Add Worksheets("Modelling Categories").Range("D7").Value, Worksheets("Modelling Categories").Range("E7").Value
        .Add Worksheets("Modelling Categories").Range("D8").Value, Worksheets("Modelling Categories").Range("E8").Value
        .Add Worksheets("Modelling Categories").Range("D9").Value, Worksheets("Modelling Categories").Range("E9").Value
        .Add Worksheets("Modelling Categories").Range("D10").Value, Worksheets("Modelling Categories").Range("E10").Value
        .Add Worksheets("Modelling Categories").Range("D11").Value, Worksheets("Modelling Categories").Range("E11").Value
        .Add Worksheets("Modelling Categories").Range("D12").Value, Worksheets("Modelling Categories").Range("E12").Value
        .Add Worksheets("Modelling Categories").Range("D13").Value, Worksheets("Modelling Categories").Range("E13").Value
        .Add Worksheets("Modelling Categories").Range("D14").Value, Worksheets("Modelling Categories").Range("E14").Value
    End With
    
    Dim Size_Cat As New Scripting.Dictionary
    With Size_Cat
        .Add Worksheets("Modelling Categories").Range("G2").Value, Worksheets("Modelling Categories").Range("H2").Value
        .Add Worksheets("Modelling Categories").Range("G3").Value, Worksheets("Modelling Categories").Range("H3").Value
        .Add Worksheets("Modelling Categories").Range("G4").Value, Worksheets("Modelling Categories").Range("H4").Value
        .Add Worksheets("Modelling Categories").Range("G5").Value, Worksheets("Modelling Categories").Range("H5").Value
        .Add Worksheets("Modelling Categories").Range("G6").Value, Worksheets("Modelling Categories").Range("H6").Value
        .Add Worksheets("Modelling Categories").Range("G7").Value, Worksheets("Modelling Categories").Range("H7").Value
        .Add Worksheets("Modelling Categories").Range("G8").Value, Worksheets("Modelling Categories").Range("H8").Value
        .Add Worksheets("Modelling Categories").Range("G9").Value, Worksheets("Modelling Categories").Range("H9").Value
        .Add Worksheets("Modelling Categories").Range("G10").Value, Worksheets("Modelling Categories").Range("H10").Value
        .Add Worksheets("Modelling Categories").Range("G11").Value, Worksheets("Modelling Categories").Range("H11").Value
        .Add Worksheets("Modelling Categories").Range("G12").Value, Worksheets("Modelling Categories").Range("H12").Value
        .Add Worksheets("Modelling Categories").Range("G13").Value, Worksheets("Modelling Categories").Range("H13").Value
        .Add Worksheets("Modelling Categories").Range("G14").Value, Worksheets("Modelling Categories").Range("H14").Value
        .Add Worksheets("Modelling Categories").Range("G15").Value, Worksheets("Modelling Categories").Range("H15").Value
        .Add Worksheets("Modelling Categories").Range("G16").Value, Worksheets("Modelling Categories").Range("H16").Value
        .Add Worksheets("Modelling Categories").Range("G17").Value, Worksheets("Modelling Categories").Range("H17").Value
    End With
    
    Dim Type_Cat As New Scripting.Dictionary
    With Type_Cat
        .Add Worksheets("Modelling Categories").Range("J2").Value, Worksheets("Modelling Categories").Range("K2").Value
        .Add Worksheets("Modelling Categories").Range("J3").Value, Worksheets("Modelling Categories").Range("K3").Value
        .Add Worksheets("Modelling Categories").Range("J4").Value, Worksheets("Modelling Categories").Range("K4").Value
        .Add Worksheets("Modelling Categories").Range("J5").Value, Worksheets("Modelling Categories").Range("K5").Value
        .Add Worksheets("Modelling Categories").Range("J6").Value, Worksheets("Modelling Categories").Range("K6").Value
        .Add Worksheets("Modelling Categories").Range("J7").Value, Worksheets("Modelling Categories").Range("K7").Value
        .Add Worksheets("Modelling Categories").Range("J8").Value, Worksheets("Modelling Categories").Range("K8").Value
    End With
'-------------------------------------------------------
    
    'Set number of records in sheet
    numrecords = UBound(valves)
    'initialize check to identify end of file
    j = 0
    For i = 1 To numrecords
        'If record is empty, increment check and continue reviewing sheet
        If IsEmpty(valves(i, 1)) Then
            'Increment j
            j = j + 1
            'Escapes from for loop when two empty rows found
            If j > 1 Then
                numrecords = i - j - 1
                Exit For
            End If
        End If
    Next i
' ------------------------------------------------------
    'Initialize variables for containing each record's information
    Dim curoperateLoF As Double
    Dim cur_Leak_LoF As Double
    Dim curValve_Postion As String
    Dim curLoc_class As String
    Dim curValve_crit As String
    Dim curInspectCost As Double
    Dim Last_Inspection As Double
    Dim curValve_Place As String
    Dim curValve_vault As String

    'Intermediate variables for Use
    Dim curCritical_class As String
    Dim optConstrained_interval As Double
    Dim optInterval As Double
    Dim optConstrained_cost As Double
    Dim curoptCost As Double
    Dim curInterval As Double
    Dim anntimefailed_ops As Double
    Dim anntimefailed_leak As Double
    Dim curCost As Double
    Dim curannLeakRisk As Double
    Dim curannOpsRisk As Double
    Dim curopsCoF As Double
    Dim curleakCoF As Double
    Dim max_iterations As Double
    Dim curconstopsPoF As Double
    Dim curconstleakPoF As Double
    Dim cur_ann_leak_risk(10) As Double
    Dim cur_ann_ops_risk(10) As Double
    
    'Initialize variables for output storage
    ReDim curValve_Use(numrecords) As String
    ReDim OutputArray(numrecords, 35) As Variant
    ReDim Optimal_Survey_Interval(numrecords) As Double
    ReDim Optimal_Constrained_Interval(numrecords) As Double
    ReDim Optimal_Annual_Cost(numrecords) As Double
    ReDim Optimal_Constrained_Annual_Cost(numrecords) As Double
    ReDim Next_Inspection(numrecords) As Double
    ReDim Constraint_Criteria(numrecords) As String
    ReDim Yr1_Annual_Cost(numrecords) As Double
    ReDim Yr2_Annual_Cost(numrecords) As Double
    ReDim Yr3_Annual_Cost(numrecords) As Double
    ReDim Yr4_Annual_Cost(numrecords) As Double
    ReDim Yr5_Annual_Cost(numrecords) As Double
    ReDim Yr6_Annual_Cost(numrecords) As Double
    ReDim Yr7_Annual_Cost(numrecords) As Double
    ReDim Yr8_Annual_Cost(numrecords) As Double
    ReDim Yr9_Annual_Cost(numrecords) As Double
    ReDim Yr10_Annual_Cost(numrecords) As Double
    ReDim Winter_Flag(numrecords) As String
    ReDim Min_reliable_optimal(numrecords) As Double
    ReDim Min_reliable_constrained(numrecords) As Double
    ReDim Ann_Inspect_Cost(numrecords) As Double
    ReDim Ann_Leak_Risk(numrecords) As Double
    ReDim Ann_Ops_Risk(numrecords) As Double
    ReDim Ann_Const_Leak_Risk(numrecords) As Double
    ReDim Ann_Const_Ops_Risk(numrecords) As Double
    ReDim Ann_Unconst_Inspect_Cost(numrecords) As Double
    ReDim Ann_Leak_LoF(numrecords) As Double
    ReDim Ann_Ops_LoF(numrecords) As Double
    
    
' ------------------------CALCULATIONS------------------------------

    For i = 2 To numrecords + 1
        'read in valve asset data for current valve
        curValve_crit = valves(i, 11)
        curValve_Postion = valves(i, 5)
        curValve_Place = valves(i, 7)
        curValve_vault = valves(i, 12)
        If valves(i, 13) = "YES" Then
            curValve_Use(i - 1) = "Heat Area"
        Else
            curValve_Use(i - 1) = "Isolation"
        End If
        curLoc_class = valves(i, 6)
        Last_Inspection = valves(i, 25)
        
        'Set inspection cost for current valve based on valve status
        '*****CHECK WHEN ACTUAL FIELD VALVES AVAILABLE*****
        If curValve_vault = "YES" Then
            curInspectCost = inspection_cost_Vault
        ElseIf curValve_Place = "ABOVE" Then
            curInspectCost = inspection_cost_AG
        Else
            curInspectCost = inspection_cost_BG
        End If

        
        'Determine which class current valve falls under for constraints and criticality measures
        '*****CHECK WHEN ACTUAL FIELD VALVES AVAILABLE*****

        If curValve_crit = "URBAN-CRIT" Then
            If curValve_vault = "YES" Then
                curCritical_class = "Urban Crit - Vault"
            ElseIf curValve_Place = "ABOVE" Then
                curCritical_class = "Urban Crit - AG"
            Else
                curCritical_class = "Urban Crit - BG"
            End If
        ElseIf curValve_crit = "URBAN" Then
            If curValve_vault = "YES" Then
                curCritical_class = "Urban - Vault"
            ElseIf curValve_Place = "ABOVE" Then
                curCritical_class = "Urban - AG"
            Else
                curCritical_class = "Urban - BG"
            End If
        ElseIf curValve_crit = "RURAL" Then
            If curValve_vault = "YES" Then
                curCritical_class = "Rural - Vault"
            ElseIf curValve_Place = "ABOVE" Then
                curCritical_class = "Rural - AG"
            Else
                curCritical_class = "Rural - BG"
            End If
        End If
        
        'record minimum inspection interval
        Constraint_Criteria(i - 1) = "<= " & min_survey_int(curCritical_class) & " Years"

        'Determine consequence cost of failures based on valve type, position, and current criticalist escalation factors
        If curValve_Postion = "OPEN" Then
            If curValve_Use(i - 1) = "Heat Area" Then
                curopsCoF = HA_ops_fail_cost * HA_ops_incident_rate * funct_incident_factors_ha(curCritical_class) _
                                                                    * funct_consequence_factors_ha(curCritical_class) _
                                                                    * funct_failure_factors_ha(curCritical_class) _
                                                                    * funct_scaling_factors_ha(curCritical_class)
                curleakCoF = ha_leak_cost * ha_leak_incident_rate * leak_incident_factors_ha(curCritical_class) _
                                                                  * leak_consequence_factors_ha(curCritical_class) _
                                                                  * funct_scaling_factors_ha(curCritical_class)
            Else
                curopsCoF = ISO_ops_fail_cost * ISO_ops_incident_rate * funct_incident_factors_iso(curCritical_class) _
                                                                      * funct_consequence_factors_Iso(curCritical_class) _
                                                                      * funct_failure_factors_Iso(curCritical_class) _
                                                                      * funct_scaling_factors_Iso(curCritical_class)
                curleakCoF = iso_leak_cost * iso_leak_incident_rate * leak_incident_factors_iso(curCritical_class) _
                                                                    * leak_consequence_factors_iso(curCritical_class) _
                                                                    * funct_scaling_factors_Iso(curCritical_class)
            End If
        'Note leak consequences are the same for open and closed valves, only functional failure is different
        Else
            If curValve_Use(i - 1) = "Heat Area" Then
                curopsCoF = HA_open_incident_rate * HA_open_fail_cost _
                                                  * funct_opening_consequence_ha(curCritical_class) _
                                                  * funct_scaling_factors_ha(curCritical_class)
                curleakCoF = ha_leak_cost * ha_leak_incident_rate * leak_incident_factors_ha(curCritical_class) _
                                                                  * leak_consequence_factors_ha(curCritical_class) _
                                                                  * funct_scaling_factors_ha(curCritical_class)
            Else
                curopsCoF = ISO_open_incident_rate * ISO_open_fail_cost _
                                                   * funct_opening_consequence_Iso(curCritical_class) _
                                                   * funct_scaling_factors_Iso(curCritical_class)
                curleakCoF = iso_leak_cost * iso_leak_incident_rate * leak_incident_factors_iso(curCritical_class) _
                                                                    * leak_consequence_factors_iso(curCritical_class) _
                                                                    * funct_scaling_factors_Iso(curCritical_class)
            End If
        End If
        
   '------Run optimization to find optimal inspection interval to minimize total cost for current valve-----
        'Error check for min_survey_Iny
        If min_survey_int(curCritical_class) = 0 Then
            MsgBox "Error: Minimum Inspection Frequency cannot be left blank"
            Worksheets("Inputs").Shapes("CommandButton1").Visible = True
            Worksheets("Inputs").Shapes("CommandButton2").Visible = True
            Worksheets("Inputs").Shapes("CommandButton3").Visible = True
            Exit Sub
        End If
        
        'set current optimal inspection interval to 1
        curInterval = 1
        'first iteration initialize starting optimal intervals and costs - set to inspecting at minimum intervals as default with extreme cost
        '           as no risks yet calculated
        curoptCost = (curInspectCost / min_survey_int(curCritical_class)) _
                    + ((min_survey_int(curCritical_class) / 2) * (1000000000))
        optInterval = min_survey_int(curCritical_class)
        optConstrained_cost = curoptCost
        optConstrained_interval = min_survey_int(curCritical_class)


        'set max iterations -- LIKELY A USER INPUT - TO INCLUDE
        max_iterations = 10
        
        j = 0
        
        Do While j < max_iterations
            
            'Run Risk Calculation for given time interval for current asset
            Dim LOF_Results As Variant
            If curValve_Place = "ABOVE" Then
                LOF_Results = LOF_Model(curInterval, i, AG_Access_Model_Params, AG_Locate_Model_Params, AG_Actuate_Model_Params, AG_Leak_Model_Params, _
                                                        Loc_Code_Cat, Pressure_Cat, Size_Cat, Type_Cat)
            ElseIf curValve_Place = "BELOW" Then
                LOF_Results = LOF_Model(curInterval, i, BG_Access_Model_Params, BG_Locate_Model_Params, BG_Actuate_Model_Params, BG_Leak_Model_Params, _
                                                        Loc_Code_Cat, Pressure_Cat, Size_Cat, Type_Cat)
            'If abandoned or buried in place - assign 0 risk - should be in the tool but this will catch if they are
            Else
                LOF_Results(1) = 0
                LOF_Results(2) = 0
            End If
    
            'Read results of risk model
            curoperateLoF = LOF_Results(1)
            cur_Leak_LoF = LOF_Results(2)
            
            'Calculate time failed across  failure types
            anntimefailed_ops = curoperateLoF / 2
            anntimefailed_leak = cur_Leak_LoF / 2

            'Calculate Annual Risk (PoF*CoF) across all failure types
            curannOpsRisk = anntimefailed_ops * curopsCoF
            curannLeakRisk = anntimefailed_leak * curleakCoF
                    
            'Calculate total annual cost of inspections and consequences given current interval
            curCost = (curInspectCost / curInterval) + (curannOpsRisk + curannLeakRisk)
                              
            'Update Optimal cost if new cost is less than old optimal cost
            'if interval is less than minimum, update constrained optimal cost, if not, just update unconstrained optimal cost
            If curInterval <= min_survey_int(curCritical_class) Then
                If curCost < curoptCost Then
                    optConstrained_interval = curInterval
                    optConstrained_cost = curCost
                    curoptCost = curCost
                    optInterval = curInterval
                    Ann_Const_Leak_Risk(i - 1) = curannLeakRisk
                    Ann_Const_Ops_Risk(i - 1) = curannOpsRisk
                    Ann_Leak_Risk(i - 1) = curannLeakRisk
                    Ann_Ops_Risk(i - 1) = curannOpsRisk
                    curconstopsPoF = curoperateLoF
                    curconstleakPoF = cur_Leak_LoF
                End If
            ElseIf curCost < curoptCost Then
                curoptCost = curCost
                optInterval = curInterval
                Ann_Leak_Risk(i - 1) = curannLeakRisk
                Ann_Ops_Risk(i - 1) = curannOpsRisk
            End If
                
            'Recording results for each inspection interval attempted
            cur_ann_leak_risk(j) = curannLeakRisk
            cur_ann_ops_risk(j) = curannOpsRisk
            'Increase interval
            curInterval = curInterval + 1
            
            'Allow Excel to perform some other functions or break out of code, and update status bar
            application.StatusBar = "Progress: " & Format(i / (numrecords + 2), "Percent") & " Complete"
            DoEvents
        j = j + 1
        
        Loop
        
    '---------------------------------------------------------------------
        'Make winter flag
        If Do_Winter = "Yes" Then
            If valves(i, 26) >= Winter_Cutoff Then
                Winter_Flag(i - 1) = "Yes"
            Else
                Winter_Flag(i - 1) = "No"
            End If
        Else
            Winter_Flag(i - 1) = "No"
        End If
            
    '---------------------------------------------------------------------

        'update output arrays
        Optimal_Survey_Interval(i - 1) = optInterval
        Optimal_Constrained_Interval(i - 1) = optConstrained_interval
        Optimal_Annual_Cost(i - 1) = curoptCost
        Optimal_Constrained_Annual_Cost(i - 1) = optConstrained_cost
        Ann_Inspect_Cost(i - 1) = (curInspectCost / optConstrained_interval)
        Ann_Unconst_Inspect_Cost(i - 1) = (curInspectCost / optInterval)
        
        'Calculate average annual cost for 1-10 year inspection interval
        Yr1_Annual_Cost(i - 1) = (curInspectCost / 1) + (cur_ann_ops_risk(0) + cur_ann_leak_risk(0))
        Yr2_Annual_Cost(i - 1) = (curInspectCost / 2) + (cur_ann_ops_risk(1) + cur_ann_leak_risk(1))
        Yr3_Annual_Cost(i - 1) = (curInspectCost / 3) + (cur_ann_ops_risk(2) + cur_ann_leak_risk(2))
        Yr4_Annual_Cost(i - 1) = (curInspectCost / 4) + (cur_ann_ops_risk(3) + cur_ann_leak_risk(3))
        Yr5_Annual_Cost(i - 1) = (curInspectCost / 5) + (cur_ann_ops_risk(4) + cur_ann_leak_risk(4))
        Yr6_Annual_Cost(i - 1) = (curInspectCost / 6) + (cur_ann_ops_risk(5) + cur_ann_leak_risk(5))
        Yr7_Annual_Cost(i - 1) = (curInspectCost / 7) + (cur_ann_ops_risk(6) + cur_ann_leak_risk(6))
        Yr8_Annual_Cost(i - 1) = (curInspectCost / 8) + (cur_ann_ops_risk(7) + cur_ann_leak_risk(7))
        Yr9_Annual_Cost(i - 1) = (curInspectCost / 9) + (cur_ann_ops_risk(8) + cur_ann_leak_risk(8))
        Yr10_Annual_Cost(i - 1) = (curInspectCost / 10) + (cur_ann_ops_risk(9) + cur_ann_leak_risk(9))

        'Calculate Next Inspection Year
        If Last_Inspection < 1 Then
            Next_Inspection(i - 1) = Year(Date)
        ElseIf Last_Inspection + optConstrained_interval < Year(Date) Then
            Next_Inspection(i - 1) = Year(Date)
        Else
            Next_Inspection(i - 1) = Last_Inspection + optConstrained_interval
        End If
    
    Next i

'output information to output array and send to output worksheet
    For i = 2 To numrecords + 1
        'Copy calculated values to OutputArray for easy output to screen
        OutputArray(i - 1, 0) = valves(i, 1)
        OutputArray(i - 1, 1) = "Valve Inspection - " & valves(i, 1)
        OutputArray(i - 1, 2) = valves(i, 18)
        OutputArray(i - 1, 3) = valves(i, 17)
        OutputArray(i - 1, 4) = valves(i, 11)
        OutputArray(i - 1, 5) = curValve_Use(i - 1)
        OutputArray(i - 1, 6) = Optimal_Survey_Interval(i - 1)
        OutputArray(i - 1, 7) = Optimal_Annual_Cost(i - 1)
        OutputArray(i - 1, 8) = Optimal_Constrained_Interval(i - 1)
        OutputArray(i - 1, 9) = Optimal_Constrained_Annual_Cost(i - 1)
        OutputArray(i - 1, 10) = Next_Inspection(i - 1)
        OutputArray(i - 1, 11) = Constraint_Criteria(i - 1)
        OutputArray(i - 1, 12) = Yr1_Annual_Cost(i - 1)
        OutputArray(i - 1, 13) = Yr2_Annual_Cost(i - 1)
        OutputArray(i - 1, 14) = Yr3_Annual_Cost(i - 1)
        OutputArray(i - 1, 15) = Yr4_Annual_Cost(i - 1)
        OutputArray(i - 1, 16) = Yr5_Annual_Cost(i - 1)
        OutputArray(i - 1, 17) = Yr6_Annual_Cost(i - 1)
        OutputArray(i - 1, 18) = Yr7_Annual_Cost(i - 1)
        OutputArray(i - 1, 19) = Yr8_Annual_Cost(i - 1)
        OutputArray(i - 1, 20) = Yr9_Annual_Cost(i - 1)
        OutputArray(i - 1, 21) = Yr10_Annual_Cost(i - 1)
        OutputArray(i - 1, 22) = Winter_Flag(i - 1)
        OutputArray(i - 1, 23) = valves(i, 27)
        OutputArray(i - 1, 24) = valves(i, 28)
        OutputArray(i - 1, 25) = valves(i, 29)
        OutputArray(i - 1, 26) = valves(i, 30)
        OutputArray(i - 1, 27) = Ann_Inspect_Cost(i - 1)
        OutputArray(i - 1, 28) = Ann_Const_Leak_Risk(i - 1)
        OutputArray(i - 1, 29) = Ann_Const_Ops_Risk(i - 1)
        OutputArray(i - 1, 30) = Ann_Unconst_Inspect_Cost(i - 1)
        OutputArray(i - 1, 31) = Ann_Leak_Risk(i - 1)
        OutputArray(i - 1, 32) = Ann_Ops_Risk(i - 1)
    Next i
    
    'Set OutputArray headers and output file to spreadsheet
    OutputArray(0, 0) = "Asset Number"
    OutputArray(0, 1) = "Description"
    OutputArray(0, 2) = "PMNUM"
    OutputArray(0, 3) = "Asset Tag (formerly Valve ID)"
    OutputArray(0, 4) = "Criticality Class"
    OutputArray(0, 5) = "Valve Use"
    OutputArray(0, 6) = "Unconstrained Optimal Interval (Year)"
    OutputArray(0, 7) = "Unconstrained Optimal Annual Cost($)"
    OutputArray(0, 8) = "Constrained Optimal Interval (Year)"
    OutputArray(0, 9) = "Constrained Optimal Annual Cost($)"
    OutputArray(0, 10) = "Next Inspection"
    OutputArray(0, 11) = "Constraint Criteria"
    OutputArray(0, 12) = "1 Year Interval - Annual Cost"
    OutputArray(0, 13) = "2 Year Interval - Annual Cost"
    OutputArray(0, 14) = "3 Year Interval - Annual Cost"
    OutputArray(0, 15) = "4 Year Interval - Annual Cost"
    OutputArray(0, 16) = "5 Year Interval - Annual Cost"
    OutputArray(0, 17) = "6 Year Interval - Annual Cost"
    OutputArray(0, 18) = "7 Year Interval - Annual Cost"
    OutputArray(0, 19) = "8 Year Interval - Annual Cost"
    OutputArray(0, 20) = "9 Year Interval - Annual Cost"
    OutputArray(0, 21) = "10 Year Interval - Annual Cost"
    OutputArray(0, 22) = "Winter Flag"
    OutputArray(0, 23) = "Accessibility History"
    OutputArray(0, 24) = "Locatability History"
    OutputArray(0, 25) = "Actuation History"
    OutputArray(0, 26) = "Leak History"
    OutputArray(0, 27) = "Annual Inspection Cost (Constrained)($)"
    OutputArray(0, 28) = "Annual Leak Risk (Constrained)($)"
    OutputArray(0, 29) = "Annual Operational Risk (Constrained)($)"
    OutputArray(0, 30) = "Annual Inspection Cost(Unconstrained)($)"
    OutputArray(0, 31) = "Annual Leak Risk(Unconstrained)($)"
    OutputArray(0, 32) = "Annual Operational Risk(Unconstrained)($)"

    
    Worksheets("Results_List").Range("A6:AJ" & numrecords + 6).Value = OutputArray
    
    'be 99% complete until levelling complete
    application.StatusBar = "Progress: " & Format(0.9999, "Percent") & " Complete"
    
    'CALL LEVELLING SUB:
    Call Inspection_Levelling
    
    'Finally Complete statusbar
    application.StatusBar = "Progress: " & Format(1, "Percent") & " Complete"
    application.StatusBar = False
' -------------------------------CALCULATE SUMMARY STATISTICS FOR RESULTS SHEET-----------------------
    
    'Show buttons
    Worksheets("Inputs").Shapes("CommandButton1").Visible = True
    Worksheets("Inputs").Shapes("CommandButton2").Visible = True
    Worksheets("Inputs").Shapes("CommandButton3").Visible = True
       
    
    'Set Date and Time Stamp
    Worksheets("Inputs").Range("G7").Value = Now
    
    'Message of successful completion
    MsgBox "Optimization Routine Completed Successfully."
    
    Exit Sub
    
ErrHandler:
    Dim lContinue As Long
    If Err.Number = 18 Then
        lContinue = MsgBox(prompt:= _
          "Do you want to stop the optimization routine?" & vbCrLf, _
          Buttons:=vbYesNo)
        If lContinue = vbNo Then
            Resume
        Else
            Worksheets("Inputs").Shapes("CommandButton1").Visible = True
            Worksheets("Inputs").Shapes("CommandButton2").Visible = True
            Worksheets("Inputs").Shapes("CommandButton3").Visible = True
            MsgBox ("Routine ended at user request")
            Exit Sub
        End If
    Else
        MsgBox ("An error has occurred. Routine not completed." _
            & vbNewLine & vbNewLine & "Error Code " & Err & ": " & Error(Err))
        application.Workbooks(curbook).Worksheets("Inputs").Shapes("CommandButton1").Visible = True
        application.Workbooks(curbook).Worksheets("Inputs").Shapes("CommandButton2").Visible = True
        application.Workbooks(curbook).Worksheets("Inputs").Shapes("CommandButton3").Visible = True
        Exit Sub
    End If

End Sub

