Option Explicit
Implements IView

Private WithEvents mViewModel As SomeViewModel
Private Type TState
    IsCancelled As Boolean
End Type
Private This As TState

Private Sub cmdCancel_Click()
    OnCancel
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set mViewModel = ViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function