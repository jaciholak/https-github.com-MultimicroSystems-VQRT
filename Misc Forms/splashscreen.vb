Public Class splashscreen
    Private Sub splashscreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Label4.Text = "1841 Montreal Rd, Suite 105, Tucker, GA 30084" & vbCrLf & "(404) - 296 - 8966" & vbCrLf & "www.multimicrosystems.com" & vbCrLf & "Copyright (C) 1984-" & Now.Year & ". All rights reserved."
    End Sub '02-06-11 Jtc 01-20-11 Fix Year

    Private Sub Label4_Click(sender As System.Object, e As System.EventArgs) Handles Label4.Click

    End Sub
End Class