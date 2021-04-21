VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Opens the AddToDo Userform'

Sub openAddToDo1()

popup_addtodo.Show

'Sets cursor on first textbox'

popup_addtodo.ToDoName.SetFocus

End Sub

'Opens the AddTasks Userform'

Sub openAddTasks()

popup_addtask.Show

'Sets cursor on first textbox'

popup_addtask.TaskClass.SetFocus

End Sub

Private Sub DeleteToDo_Click()
    ActiveCell.EntireRow = ""
End Sub

Private Sub StatusUpdate_Click()

If Range("H" & (ActiveCell.Row)).Value = "Incomplete" Then
    Range("H" & (ActiveCell.Row)).Value = "Complete"

ElseIf Range("H" & (ActiveCell.Row)).Value = "Complete" Then
    Range("H" & (ActiveCell.Row)).Value = "Incomplete"
 
End If

End Sub

Sub HelpToDo()

help_todo.Show

End Sub

Sub HelpTask()

help_task.Show

End Sub
