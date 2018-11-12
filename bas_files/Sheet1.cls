VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox1_Click()
    'Allows VBA to edit sheets while keeping interface protected
    Worksheets("Valve_Optimization_Controls").Unprotect
    Worksheets("Valve_Optimization_Controls").Protect , userinterfaceonly:=True
    If Worksheets("Valve_Optimization_Controls").Value = "No" Then
        Worksheets("Valve_Optimization_Controls").Value = "Yes"
    Else
        Worksheets("Valve_Optimization_Controls").Value = "No"
    End If
        
End Sub

Sub AreYouSure()

Dim Sure As Integer

Sure = MsgBox("This may take up to 5 minutes to run. To cancel at any time, press ctrl+break and select End." _
                & vbNewLine & vbNewLine & "Progress is indicated in the bottom left corner." _
                & vbNewLine & vbNewLine & "Do you want to continue?", vbOKCancel)
If Sure = 1 Then Call JVIO_Optimization

End Sub


Private Sub CommandButton1_Click()
CommandButton1.BackColor = 13464600
Call AreYouSure
End Sub

Private Sub CommandButton2_Click()
CommandButton2.BackColor = 13464600
Call Unlock_Inputs
Call Unlock_Inputs
End Sub

Private Sub CommandButton3_Click()
CommandButton3.BackColor = 13464600
Call Reset_to_Default
End Sub

Private Sub CommandButton4_Click()
CommandButton4.BackColor = 13464600
Call Quick_Print
End Sub


