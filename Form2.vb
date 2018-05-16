Imports System.Reflection
Public Class Form2


    Private Sub Form2_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        Me.Label2.Text = "Version: " & Assembly.GetExecutingAssembly().GetName().Version.ToString
    End Sub
End Class