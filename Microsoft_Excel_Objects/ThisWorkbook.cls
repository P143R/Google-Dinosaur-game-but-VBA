VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_Open()
If ThisWorkbook.AutoSaveOn Then
    ActiveWorkbook.AutoSaveOn = False
End If
'Start function/macro is called when spacebar (" ") is pressed
    Application.OnKey (" "), "start"
    reset
End Sub
