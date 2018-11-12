Attribute VB_Name = "Module3_Reset_to_Default"
Sub Reset_to_Default()

    'Prevent Excel from updating screen while calculations run to improve runtime
    application.ScreenUpdating = False

    'Allows VBA to edit sheet while keeping interface protected
    Worksheets("Inputs").Unprotect "QS"
    Worksheets("Inputs").Protect "QS", userinterfaceonly:=True
    Worksheets("Results_List").Unprotect "QS"
    Worksheets("Results_List").Protect "QS", userinterfaceonly:=True, AllowFiltering:=True
    Worksheets("Levelled_Inspections").Unprotect "QS"
    Worksheets("Levelled_Inspections").Protect "QS", userinterfaceonly:=True, AllowFiltering:=True
    
    'Clear Date and Time Stamp
    Worksheets("Inputs").Range("G7").ClearContents
    
    'Reset User Inputs to Defaults or to blank if no default
    Worksheets("Inputs").Range("B7:C8").Value = Worksheets("Defaults").Range("B7:C8").Value
    Worksheets("Inputs").Range("B9:F10").Value = Worksheets("Defaults").Range("B9:F10").Value
    Worksheets("Inputs").Range("B14:C17").Value = Worksheets("Defaults").Range("B14:C17").Value
    Worksheets("Inputs").Range("C23:C24").Value = Worksheets("Defaults").Range("C23:C24").Value
    Worksheets("Inputs").Range("H15:H25").Value = Worksheets("Defaults").Range("H15:H25").Value
    Worksheets("Inputs").Range("C30:D31").Value = Worksheets("Defaults").Range("C30:D31").Value
    Worksheets("Inputs").Range("E30:F31").Value = Worksheets("Defaults").Range("E30:F31").Value
    Worksheets("Inputs").Range("C36:J46").Value = Worksheets("Defaults").Range("C36:J46").Value
    Worksheets("Inputs").Range("C51:D52").Value = Worksheets("Defaults").Range("C51:D52").Value
    Worksheets("Inputs").Range("C57:F67").Value = Worksheets("Defaults").Range("C57:F67").Value
    Call Unlock_Inputs
    Call Unlock_Inputs
    
    'Clear Output Sheets
    With Worksheets("Results_List")
    .Rows("7:" & .Rows.Count).ClearContents
    End With
    
    With Worksheets("Levelled_Inspections")
    .Rows("7:" & .Rows.Count).ClearContents
    End With
      
    With Worksheets("Level_Helper")
    .Range("B3:E12").ClearContents
    .Range("B15:E24").ClearContents
    End With
    
    With Worksheets("Levelled_Inspections")
    .Rows("7:" & .Rows.Count).ClearContents
    End With
    
    Worksheets("Level_Helper").Range("B3:E12").ClearContents
    Worksheets("Level_Helper").Range("B15:E24").ClearContents
    
    'Hide unlock button
    Worksheets("Inputs").Shapes("CommandButton2").Visible = False
    Worksheets("Inputs").Shapes("CommandButton1").Visible = True
    
    'Reset updating screen
    application.ScreenUpdating = True

End Sub


