Attribute VB_Name = "Module6_Levelling"
'Ensure text comparisons are done ignoring case
Option Compare Text
'Enable Option Explicit
Option Explicit

Sub Inspection_Levelling()

    'General variables for loops etc.
    Dim i As Long
    Dim j As Long
    Dim k As Long
    'Initialize values of general variables
    i = 0
    j = 0
    k = 0
    'Initialize input variables
    Dim records As Long
    Dim valvedata As Variant
    Dim results As Variant
    
    'Read in all valve and results data
    valvedata = Worksheets("Valve_List").Range("A6:Z10000").Value
    results = Worksheets("Results_List").Range("A6:AI10000").Value
        
    'Set number of records in sheet
    records = UBound(results)
    'initialize check to identify end of file
    j = 0
    For i = 1 To records
        'If record is empty, increment check and continue reviewing sheet
        If IsEmpty(results(i, 1)) Then
            'Increment j
            j = j + 1
            'Escapes from for loop when two empty rows found
            If j > 1 Then
                records = i - j - 1
                Exit For
            End If
        End If
    Next i
    
    'intialize region totals
    Dim Annual_Total_Cost As Double
    Annual_Total_Cost = 0

    
    Dim Valve_Count As Integer
    Valve_Count = 0
    
    'Initialize Dictionaries to hold annual costs for calculation for each region
    'there is almost definitely a better way to do this but time is limited
    Dim Costs_Per_Year
    Set Costs_Per_Year = CreateObject("Scripting.Dictionary")
    With Costs_Per_Year
        .Add Year(Date), 0
        .Add Year(Date) + 1, 0
        .Add Year(Date) + 2, 0
        .Add Year(Date) + 3, 0
        .Add Year(Date) + 4, 0
        .Add Year(Date) + 5, 0
        .Add Year(Date) + 6, 0
        .Add Year(Date) + 7, 0
        .Add Year(Date) + 8, 0
        .Add Year(Date) + 9, 0
    End With
        
    Dim Count_Per_Year
    Set Count_Per_Year = CreateObject("Scripting.Dictionary")
    With Count_Per_Year
        .Add Year(Date), 0
        .Add Year(Date) + 1, 0
        .Add Year(Date) + 2, 0
        .Add Year(Date) + 3, 0
        .Add Year(Date) + 4, 0
        .Add Year(Date) + 5, 0
        .Add Year(Date) + 6, 0
        .Add Year(Date) + 7, 0
        .Add Year(Date) + 8, 0
        .Add Year(Date) + 9, 0
    End With
    
    'Define variables for results and calculations
    ReDim OutputsArray(records, 35) As Variant
    ReDim LoaderArray(records, 35) As Variant
    Dim inspect_int As Double
    Dim cur_year As Double
    Dim Source As String
    Dim inspect_cost As Double
    Dim min_cost As Double
    Dim Check_Count As Double
    Dim Prev_Inspection As Double
    Dim Last_Inspection As Double
    Dim curValve_crit As String
    Dim curValve_vault As String
    Dim curValve_Place As String
    ReDim Next_Inspection(records) As Double
    ReDim curCritical_class(records) As String
    
    cur_year = Year(Date)
    Check_Count = 0
    
    'Calculate total annual inspection cost
    For i = 2 To records + 1:
        Annual_Total_Cost = Annual_Total_Cost + results(i, 28)
        Valve_Count = Valve_Count + 1
    Next i
    
    ' Main Assignment of Valves:
    For i = 2 To records + 1:
        inspect_int = results(i, 9)
        inspect_cost = results(i, 28) * inspect_int
        '***FIX SO LEVELLING DONE WITHOUT REGIONS/SOURCES******
        Source = results(i, 5)
        Prev_Inspection = valvedata(i, 25)
        'Upper bound for assigning costs to a year = average annual total cost + $1000 to give flexibility
        min_cost = Annual_Total_Cost + 1000
        
        'Clean Previous inspection data
        'If no data for prev inspection - assume last inspection was 1 inspection interval ago
        If Prev_Inspection < 1 Then
            Last_Inspection = cur_year - inspect_int
        'if last inspection was further back thatn 1 inspection interval - reassign to
        ElseIf Prev_Inspection + inspect_int < Year(Date) Then
            Last_Inspection = cur_year - inspect_int
        Else
            Last_Inspection = Prev_Inspection
        End If
        
        'Check which year (within inspection interval) has lowest costs assigned so far and assign valve to that year
        '(ties default to nearest year)
        For j = cur_year To (Last_Inspection + inspect_int):
            If Costs_Per_Year(j) < min_cost Then
                min_cost = Costs_Per_Year(j)
                Next_Inspection(i - 1) = j
            End If
        Next j
        'Starting at next inspection - add costs each inspection interval up to 10 years
        For k = Next_Inspection(i - 1) To (cur_year + 9) Step inspect_int:
            Costs_Per_Year(k) = Costs_Per_Year(k) + inspect_cost
            Count_Per_Year(k) = Count_Per_Year(k) + 1
        Next k
        Check_Count = Check_Count + 1
    
        'Assign Criticality Designation - copied from optimization routine
        curValve_crit = valvedata(i, 11)
        curValve_vault = valvedata(i, 12)
        curValve_Place = valvedata(i, 7)
                
        If curValve_crit = "URBAN-CRIT" Then
            If curValve_vault = "YES" Then
                curCritical_class(i - 1) = "Urban Crit - Vault"
            ElseIf curValve_Place = "ABOVE" Then
                curCritical_class(i - 1) = "Urban Crit - AG"
            Else
                curCritical_class(i - 1) = "Urban Crit - BG"
            End If
        ElseIf curValve_crit = "URBAN" Then
            If curValve_vault = "YES" Then
                curCritical_class(i - 1) = "Urban - Vault"
            ElseIf curValve_Place = "ABOVE" Then
                curCritical_class(i - 1) = "Urban - AG"
            Else
                curCritical_class(i - 1) = "Urban - BG"
            End If
        ElseIf curValve_crit = "RURAL" Then
            If curValve_vault = "YES" Then
                curCritical_class(i - 1) = "Rural - Vault"
            ElseIf curValve_Place = "ABOVE" Then
                curCritical_class(i - 1) = "Rural - AG"
            Else
                curCritical_class(i - 1) = "Rural - BG"
            End If
        End If
    
    Next i

    
'output information to output array and send to output worksheet
    For i = 2 To records + 1
        'Copy calculated values to OutputArray for easy output to screen
        'FIX AREAS REFERENCING RESULTS
        OutputsArray(i - 1, 0) = results(i, 1)
        OutputsArray(i - 1, 1) = results(i, 2)
        OutputsArray(i - 1, 2) = results(i, 3)
        OutputsArray(i - 1, 3) = results(i, 4)
        OutputsArray(i - 1, 4) = results(i, 6)
        OutputsArray(i - 1, 5) = results(i, 12)
        OutputsArray(i - 1, 6) = valvedata(i, 25)
        OutputsArray(i - 1, 7) = results(i, 9)
        OutputsArray(i - 1, 8) = results(i, 9) * results(i, 28)
        OutputsArray(i - 1, 9) = Next_Inspection(i - 1)
        OutputsArray(i - 1, 10) = curCritical_class(i - 1)
        OutputsArray(i - 1, 11) = results(i, 23)
        'Copy outputs to loader sheet
        LoaderArray(i - 1, 0) = results(i, 1)
        LoaderArray(i - 1, 1) = "Valve Inspection - " & results(i, 4)
        LoaderArray(i - 1, 2) = ""
        LoaderArray(i - 1, 3) = results(i, 9)
        LoaderArray(i - 1, 4) = "YEARS"
        If results(i, 23) = "yes" Then
            LoaderArray(i - 1, 5) = "WINTINSPVLV"
        Else
            LoaderArray(i - 1, 5) = "INSPVLV"
        End If
        LoaderArray(i - 1, 6) = Next_Inspection(i - 1) & "-01-15T00:00:00+00:00"
        LoaderArray(i - 1, 7) = "CAN"
        LoaderArray(i - 1, 8) = valvedata(i, 18)
        LoaderArray(i - 1, 9) = "PNL"
        LoaderArray(i - 1, 10) = "1"
        LoaderArray(i - 1, 11) = "PM"
        LoaderArray(i - 1, 12) = "APPR"
    Next i
    
    'Set OutputArray headers and output file to spreadsheet
    OutputsArray(0, 0) = "Asset Number"
    OutputsArray(0, 1) = "Description"
    OutputsArray(0, 2) = "PMNUM"
    OutputsArray(0, 3) = "Asset Tag (Formerly Valve ID)"
    OutputsArray(0, 4) = "Valve Use"
    OutputsArray(0, 5) = "Constraint Criteria"
    OutputsArray(0, 6) = "Previous Inspection"
    OutputsArray(0, 7) = "Inspection Interval (Yr)"
    OutputsArray(0, 8) = "Inspection Cost ($)"
    OutputsArray(0, 9) = "Next Inspection"
    OutputsArray(0, 10) = "Criticality Designation"
    OutputsArray(0, 11) = "Winter Recommendation"
    'Set headers for loader
    LoaderArray(0, 0) = "ASSETNUM"
    LoaderArray(0, 1) = "DESCRIPTION"
    LoaderArray(0, 2) = "FLLASTEMAILDATE"
    LoaderArray(0, 3) = "FREQUENCY"
    LoaderArray(0, 4) = "FREQUNIT"
    LoaderArray(0, 5) = "JPNUM"
    LoaderArray(0, 6) = "NEXTDATE"
    LoaderArray(0, 7) = "ORIGID"
    LoaderArray(0, 8) = "PMNUM"
    LoaderArray(0, 9) = "SITEID"
    LoaderArray(0, 10) = "USETARGETDATE"
    LoaderArray(0, 11) = "WORKTYPE"
    LoaderArray(0, 12) = "WPSTATUS"
    
    Worksheets("Levelled_Inspections").Range("A6:AJ" & records + 6).Value = OutputsArray
    Worksheets("Results Data Load").Range("A6:AJ" & records + 6).Value = LoaderArray
'Dump Summary Outputs to Helper sheet
For j = 1 To 10:
    Worksheets("Level_Helper").Range("A" & j + 2).Value = (j + cur_year - 1)
    Worksheets("Level_Helper").Range("B" & j + 2).Value = Costs_Per_Year(j + cur_year - 1)
    Worksheets("Level_Helper").Range("A" & j + 14).Value = (j + cur_year - 1)
    Worksheets("Level_Helper").Range("B" & j + 14).Value = Count_Per_Year(j + cur_year - 1)

Next j

Worksheets("Level_Helper").Range("G2").Value = Check_Count
Worksheets("Level_Helper").Range("H2").Value = Valve_Count

End Sub

