Public Class Message_Box
    Private Property MoveForm As Boolean
    Private Property MoveForm_MousePosition As Point
    Private seconds As Integer
    Public Sub New(ByVal reply_type As Integer, ByVal reply_mes As String)
        ' This call is required by the designer.
        InitializeComponent()
        If reply_type = 0 Then 'error
            lblTitle.Text = "Error"
            pbImg.Image = My.Resources._error
            lblErrorTitle.Text = "Error:"
            lblMes.Text = reply_mes
        ElseIf reply_type = 1 Then 'success
            lblTitle.Text = "Success"
            pbImg.Image = My.Resources.success
            lblErrorTitle.Text = "Success:"
            lblMes.Text = reply_mes
        End If
        Me.Show()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Protected Overrides Sub Finalize()  ' destructor
    End Sub
    Private Sub Message_Box_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp, pnlMessageBox.MouseUp
        If e.Button = MouseButtons.Left Then
            MoveForm = False
            MoveForm_MousePosition = e.Location
        End If
    End Sub
    Private Sub Message_Box_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove, pnlMessageBox.MouseMove
        If (MoveForm = True) Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If
    End Sub
    Private Sub Message_Box_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown, pnlMessageBox.MouseDown
        If e.Button = MouseButtons.Left Then
            MoveForm = True
            Me.Cursor = Cursors.Default
            MoveForm_MousePosition = e.Location
        End If
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        seconds += 1
        If seconds >= 5 Then
            Timer1.Dispose()
            Me.Close()
        End If
    End Sub
End Class