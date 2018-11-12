Attribute VB_Name = "Module2_Unlock_Inputs"
Sub Unlock_Inputs()

    'Prevent Excel from updating screen while calculations run to improve runtime
    application.ScreenUpdating = False

    'Allows VBA to edit sheet while keeping interface protected
    Worksheets("Inputs").Unprotect "QS"
    Worksheets("Inputs").Protect "QS", userinterfaceonly:=True

    'Clear Date and Time Stamp
    Worksheets("Inputs").Range("G7").ClearContents
    
    'Unlock User Inputs
    Worksheets("Inputs").Range("B7:C8").Locked = False
    Worksheets("Inputs").Range("B9:F10").Locked = False
    Worksheets("Inputs").Range("B14:C17").Locked = False
    Worksheets("Inputs").Range("C23:C24").Locked = False
    Worksheets("Inputs").Range("H15:H25").Locked = False
    Worksheets("Inputs").Range("C30:D31").Locked = False
    Worksheets("Inputs").Range("E30:F31").Locked = False
    Worksheets("Inputs").Range("C36:J46").Locked = False
    Worksheets("Inputs").Range("C51:D52").Locked = False
    Worksheets("Inputs").Range("C57:F67").Locked = False


    'Show buttons that were hidden
    Worksheets("Inputs").Shapes("CommandButton1").Visible = True
    Worksheets("Inputs").Shapes("CommandButton3").Visible = True

    'Reset updating screen
    application.ScreenUpdating = True
    
    Worksheets("Inputs").Range("B7").Select
    

End Sub
