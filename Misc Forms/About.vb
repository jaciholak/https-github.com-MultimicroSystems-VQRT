Public Class frmAbout
  
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer '03-26-04

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.Text = "Copyright 1984-" & Now.Year
    End Sub

    Private Sub C1SuperLabel1_LinkClicked(ByVal sender As Object, ByVal e As C1.Win.C1SuperTooltip.C1SuperLabelLinkClickedEventArgs)
        Dim taskid As Short = ShellExecute(Handle.ToInt32, "OPEN", "http://www.multimicrosystems.com", "", "", AppWinStyle.MaximizedFocus) '02-28-06
    End Sub
End Class