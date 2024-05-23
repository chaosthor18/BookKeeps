Imports System.Windows.Forms
Public Class ViewSampleJournal
    Private Property MoveForm As Boolean
    Private Property MoveForm_MousePosition As Point

    Public title_Tab As String = ""
    Public Sub set_Title(ByVal name As String)
        title_Tab = name
    End Sub
    Public Sub set_Image()
        lblTitle.Text = title_Tab
        If (lblTitle.Text = "Sales on Credit") Then
            pbSample.Image = My.Resources.salesoncredit
        ElseIf (lblTitle.Text = "Purchase on Credit") Then
            pbSample.Image = My.Resources.purchaseoncredit
        ElseIf (lblTitle.Text = "Recording quarterly or year end inventory") Then
            pbSample.Image = My.Resources.yearendinventory
        ElseIf (lblTitle.Text = "Recording Of depreciation On assets") Then
            pbSample.Image = My.Resources.depreciation
        End If
    End Sub
    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Me.Close()
    End Sub

    Private Sub ViewSampleJournal_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown, pnlControl.MouseDown
        If (e.Button = MouseButtons.Left) Then
            MoveForm = True
            Me.Cursor = Cursors.Default
            MoveForm_MousePosition = e.Location
        End If
    End Sub

    Private Sub ViewSampleJournal_MouseUp(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp, pnlControl.MouseUp
        If (e.Button = MouseButtons.Left) Then
            MoveForm = False
            MoveForm_MousePosition = e.Location
        End If
    End Sub

    Private Sub pnlControl_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseUp, pnlControl.MouseMove
        If (MoveForm = True) Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If
    End Sub
End Class
