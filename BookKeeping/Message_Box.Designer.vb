<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Message_Box
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Message_Box))
        Me.btnOk = New System.Windows.Forms.Button()
        Me.pnlMessageBox = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblErrorTitle = New System.Windows.Forms.Label()
        Me.lblMes = New System.Windows.Forms.Label()
        Me.pbImg = New System.Windows.Forms.PictureBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.pnlMessageBox.SuspendLayout()
        CType(Me.pbImg, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.Location = New System.Drawing.Point(261, 162)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 47)
        Me.btnOk.TabIndex = 24
        Me.btnOk.Text = "Ok"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'pnlMessageBox
        '
        Me.pnlMessageBox.BackColor = System.Drawing.Color.Black
        Me.pnlMessageBox.Controls.Add(Me.lblTitle)
        Me.pnlMessageBox.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMessageBox.Location = New System.Drawing.Point(0, 0)
        Me.pnlMessageBox.Name = "pnlMessageBox"
        Me.pnlMessageBox.Size = New System.Drawing.Size(348, 35)
        Me.pnlMessageBox.TabIndex = 25
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(95, 18)
        Me.lblTitle.TabIndex = 8
        Me.lblTitle.Text = "BookKeeps"
        '
        'lblErrorTitle
        '
        Me.lblErrorTitle.AutoSize = True
        Me.lblErrorTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErrorTitle.Location = New System.Drawing.Point(111, 74)
        Me.lblErrorTitle.Name = "lblErrorTitle"
        Me.lblErrorTitle.Size = New System.Drawing.Size(89, 18)
        Me.lblErrorTitle.TabIndex = 27
        Me.lblErrorTitle.Text = "Error Title:"
        '
        'lblMes
        '
        Me.lblMes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMes.Location = New System.Drawing.Point(111, 92)
        Me.lblMes.Name = "lblMes"
        Me.lblMes.Size = New System.Drawing.Size(237, 52)
        Me.lblMes.TabIndex = 28
        Me.lblMes.Text = "Please check if it is a number"
        '
        'pbImg
        '
        Me.pbImg.Image = CType(resources.GetObject("pbImg.Image"), System.Drawing.Image)
        Me.pbImg.Location = New System.Drawing.Point(15, 48)
        Me.pbImg.Name = "pbImg"
        Me.pbImg.Size = New System.Drawing.Size(90, 96)
        Me.pbImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.pbImg.TabIndex = 26
        Me.pbImg.TabStop = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Message_Box
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(348, 212)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblMes)
        Me.Controls.Add(Me.lblErrorTitle)
        Me.Controls.Add(Me.pbImg)
        Me.Controls.Add(Me.pnlMessageBox)
        Me.Controls.Add(Me.btnOk)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Message_Box"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Message_Box"
        Me.pnlMessageBox.ResumeLayout(False)
        Me.pnlMessageBox.PerformLayout()
        CType(Me.pbImg, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnOk As Button
    Friend WithEvents pnlMessageBox As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents pbImg As PictureBox
    Friend WithEvents lblErrorTitle As Label
    Friend WithEvents lblMes As Label
    Friend WithEvents Timer1 As Timer
End Class
