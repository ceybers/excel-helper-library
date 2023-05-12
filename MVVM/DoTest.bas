Option Explicit

Public Sub DoTest()
    Dim ViewModel As IViewModel
    Set ViewModel = New SomeViewModel
    ViewModel.Load 
    
    Dim View As IView
    Set View = New View
    
    If View.ShowDialog(ViewModel) Then
        Debug.Print "View.ShowDialog(ViewModel) returned True"
    Else
        Debug.Print "View.ShowDialog(ViewModel) returned False"
    End If
End Sub

