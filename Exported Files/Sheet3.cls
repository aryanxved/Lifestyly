VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub openAddGrades()

popup_addgrade.Show

popup_addgrade.TextBox1.SetFocus

End Sub


Private Sub DeleteGrade_Click()
ActiveCell.EntireRow = ""
End Sub
