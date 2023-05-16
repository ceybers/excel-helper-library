Option Explicit
Implements IView

Private WithEvents mViewModel As SomeViewModel
Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdCancel_Click()
    ' Remember to set the Cancel button's property "Cancel" to true. This lets us press the Escape key to close the dialog.
    OnCancel
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function