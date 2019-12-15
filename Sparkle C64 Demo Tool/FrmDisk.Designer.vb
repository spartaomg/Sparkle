<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmDisk
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
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmDisk))
		Me.lbl = New System.Windows.Forms.Label()
		Me.pbx = New System.Windows.Forms.PictureBox()
		CType(Me.pbx, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'lbl
		'
		Me.lbl.AutoSize = True
		Me.lbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbl.Location = New System.Drawing.Point(58, 23)
		Me.lbl.Name = "lbl"
		Me.lbl.Size = New System.Drawing.Size(177, 24)
		Me.lbl.TabIndex = 0
		Me.lbl.Text = "Sparkle is working..."
		'
		'pbx
		'
		Me.pbx.BackgroundImage = CType(resources.GetObject("pbx.BackgroundImage"), System.Drawing.Image)
		Me.pbx.Location = New System.Drawing.Point(7, 11)
		Me.pbx.Name = "pbx"
		Me.pbx.Size = New System.Drawing.Size(48, 48)
		Me.pbx.TabIndex = 1
		Me.pbx.TabStop = False
		'
		'FrmDisk
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(242, 72)
		Me.Controls.Add(Me.pbx)
		Me.Controls.Add(Me.lbl)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Name = "FrmDisk"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		CType(Me.pbx, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents lbl As Label
    Friend WithEvents pbx As PictureBox
End Class
