Attribute VB_Name = "Module4_Quick_Print"
Sub Quick_Print()

    'Prints optimization controls sheet
    Worksheets("Inputs").PrintOut
    Worksheets("Results_Summary").PrintOut
    Worksheets("Levelled_Summary").PrintOut
    
End Sub

