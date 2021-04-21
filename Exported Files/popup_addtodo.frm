VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} popup_addtodo 
   Caption         =   "Add To-Do List Entry"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "popup_addtodo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "popup_addtodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm1_Click()

End Sub

Private Sub EnterButton_Click()
  Set WB = ThisWorkbook
  Set TD = WB.Worksheets("Scheduling")
  RowNum = 21

If ToDoName.Value = "" Or ToDoCourse.Value = "" Or ToDoDate.Value = "" Then
    If MsgBox("This form contains incomplete entrys, Are you sure you would like to continue?", vbQuestion + vbYesNo) <> vbYes Then
    Exit Sub
    End If
End If

Do While (TD.Cells(RowNum, "A") <> "")
    RowNum = RowNum + 1

Loop

TD.Cells(RowNum, "F") = ToDoName.Text
TD.Cells(RowNum, "G") = ToDoCourse.Text
TD.Cells(RowNum, "H") = ToDoDate.Text

'Code that does the same as above - troubleshooting'

' Set WB = ThisWorkbook
'    Set TD = WB.Worksheets("Scheduling")
'    Dim RowNum As Long
'    RowNum = 13
'
'    If (ToDoName.Value <> "") Then
'
'        Do While (TD.Cells(RowNum, "A") <> "")
'
'                    RowNum = RowNum + 1
'
'        Loop
'
'        TD.Cells(RowNum, "F") = ToDoName.Value
'
'    End If
'
'    If (ToDoName.Value <> "") Then
'
'        Do While (TD.Cells(RowNum, "A") <> "")
'
'                    RowNum = RowNum + 1
'
'
'        Loop
'
'        TD.Cells(RowNum, "G") = ToDoName.Value
'
'    End If
'    If (ToDoName.Value <> "") Then
'
'        Do While (TD.Cells(RowNum, "A") <> "")
'
'                    RowNum = RowNum + 1
'
'        Loop
'
'        TD.Cells(RowNum, "H") = ToDoName.Value
'
'    End If


'Set WB = ThisWorkbook
'   Dim ERow As Long
'  Set TD = WB.Worksheets("Scheduling")
'If ToDoName.Value = "" Or ToDoCourse.Value = "" Or ToDoDate.Value = "" Then
'    If MsgBox("This form contains incomplete entrys, Are you sure you would like to continue?", vbQuestion + vbYesNo) <> vbYes Then
'    Exit Sub
'    End If
'End If
'
'ERow = TD.Cells(Rows.Count, 13).End(xlUp).Offset(1, 0).Row
'TD.Cells(ERow, "F") = ToDoName.Value
'
'ERow = TD.Cells(Rows.Count, 13).End(xlUp).Offset(1, 0).Row
'TD.Cells(ERow, "G") = ToDoCourse.Value
'
'ERow = TD.Cells(Rows.Count, 13).End(xlUp).Offset(1, 0).Row
'TD.Cells(ERow, "H") = ToDoDate.Value

'Call emptyForm'

Unload Me

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub


Private Sub ToDoName_Change()

End Sub

Private Sub UserForm_Click()

End Sub
