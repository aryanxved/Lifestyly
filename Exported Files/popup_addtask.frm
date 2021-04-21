VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} popup_addtask 
   Caption         =   "Add Task"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "popup_addtask.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "popup_addtask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Click()

End Sub


Private Sub EnterButton1_Click()

Set WB = ThisWorkbook
  Set TD = WB.Worksheets("Scheduling")
  RowNum = 40

If TaskClass.Value = "" Or TaskName.Value = "" Then
    If MsgBox("This form contains incomplete entrys, Are you sure you would like to continue?", vbQuestion + vbYesNo) <> vbYes Then
    Exit Sub
    End If
End If

Do While (TD.Cells(RowNum, "A") <> "")
    RowNum = RowNum + 1

Loop

TD.Cells(RowNum, "F") = TaskClass.Text
TD.Cells(RowNum, "G") = TaskName.Text
TD.Cells(RowNum, "H") = "Incomplete"


'Private Sub EnterButton1_Click()
'  Set WB = ThisWorkbook
'   Dim NextRow As Long
'  Set TD = WB.Worksheets("Scheduling")
'If TaskClass.Value = "" Or TaskName.Value = "" Then
'    If MsgBox("This form contains incomplete entrys, Are you sure you would like to continue?", vbQuestion + vbYesNo) <> vbYes Then
'    Exit Sub
'    End If
'End If
'ERow = TD.Cells(Rows.Count, 33).End(xlUp).Offset(1, 0).Row
'TD.Cells(ERow, "S") = TaskClass.Value
'TD.Cells(ERow, "U") = TaskName.Value
'
'Call emptyForm

Unload Me

End Sub

Private Sub Label1_Click()

End Sub


Sub emptyForm()

TaskClass.Value = ""
TaskName.Value = ""

popup_addtask.TaskClass.SetFocus


End Sub


