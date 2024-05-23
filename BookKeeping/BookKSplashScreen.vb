Public Class BookKSplashScreen
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        pnlLoad.Width += 3
        If (pnlLoad.Width >= 540) Then
            Try
                Dim xl As Object = CreateObject("Excel.Application")
                Timer1.Dispose()
                frmMain.Show()
                Me.Close()
            Catch
                Dim msgbox1 As Message_Box = New Message_Box(0, "Please Install Microsoft Office")
                Me.Close()
            End Try
        End If
    End Sub
End Class