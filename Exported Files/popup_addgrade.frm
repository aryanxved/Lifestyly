VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} popup_addgrade 
   Caption         =   "Add GradeBook Entry"
   ClientHeight    =   7230
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11260
   OleObjectBlob   =   "popup_addgrade.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "popup_addgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub EnterButton3_Click()
  Set WB = ThisWorkbook
   Dim ERow As Long
  Set TD = WB.Worksheets("Grades")
If TextBox1.Value = "" Or TextBox2.Value = "" Or TextBox3.Value = "" Or TextBox4.Value = "" Or TextBox5.Value = "" Or TextBox6.Value = "" Or TextBox7.Value = "" Or TextBox8.Value = "" Then
    If MsgBox("This form contains incomplete entrys, Are you sure you would like to continue?", vbQuestion + vbYesNo) <> vbYes Then
    Exit Sub
    End If
End If
ERow = TD.Cells(Rows.count, 10).End(xlUp).Offset(1, 0).Row
TD.Cells(ERow, "A") = TextBox1.Value
TD.Cells(ERow, "D") = TextBox2.Value
TD.Cells(ERow, "G") = TextBox3.Value
TD.Cells(ERow, "J") = TextBox4.Value
TD.Cells(ERow, "N") = TextBox5.Value
TD.Cells(ERow, "R") = TextBox6.Value
TD.Cells(ERow, "U") = TextBox7.Value
TD.Cells(ERow, "X") = TextBox8.Value

Call emptyForm

End Sub

Sub emptyForm()

TextBox1.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""
TextBox4.Value = ""
TextBox5.Value = ""
TextBox6.Value = ""
TextBox7.Value = ""
TextBox8.Value = ""

popup_addgrade.TextBox1.SetFocus


End Sub


Private Sub UserForm_Click()

End Sub
