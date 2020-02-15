<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmEditor
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
		Me.PnlPacker = New System.Windows.Forms.Panel()
		Me.OptBetter = New System.Windows.Forms.RadioButton()
		Me.OptFaster = New System.Windows.Forms.RadioButton()
		Me.lblPacker = New System.Windows.Forms.Label()
		Me.ChkToolTips = New System.Windows.Forms.CheckBox()
		Me.ChkExpand = New System.Windows.Forms.CheckBox()
		Me.BtnPartDown = New System.Windows.Forms.Button()
		Me.BtnPartUp = New System.Windows.Forms.Button()
		Me.BtnFileDown = New System.Windows.Forms.Button()
		Me.BtnFileUp = New System.Windows.Forms.Button()
		Me.txtEdit = New System.Windows.Forms.TextBox()
		Me.Label10 = New System.Windows.Forms.Label()
		Me.BtnNew = New System.Windows.Forms.Button()
		Me.BtnCancel = New System.Windows.Forms.Button()
		Me.BtnOK = New System.Windows.Forms.Button()
		Me.BtnSave = New System.Windows.Forms.Button()
		Me.BtnLoad = New System.Windows.Forms.Button()
		Me.TV = New System.Windows.Forms.TreeView()
		Me.strip = New System.Windows.Forms.StatusStrip()
		Me.tssLabel = New System.Windows.Forms.ToolStripStatusLabel()
		Me.TssDisk = New System.Windows.Forms.ToolStripStatusLabel()
		Me.PnlPacker.SuspendLayout()
		Me.strip.SuspendLayout()
		Me.SuspendLayout()
		'
		'PnlPacker
		'
		Me.PnlPacker.Controls.Add(Me.OptBetter)
		Me.PnlPacker.Controls.Add(Me.OptFaster)
		Me.PnlPacker.Controls.Add(Me.lblPacker)
		Me.PnlPacker.Location = New System.Drawing.Point(666, 283)
		Me.PnlPacker.Name = "PnlPacker"
		Me.PnlPacker.Size = New System.Drawing.Size(108, 71)
		Me.PnlPacker.TabIndex = 125
		'
		'OptBetter
		'
		Me.OptBetter.AutoSize = True
		Me.OptBetter.Location = New System.Drawing.Point(9, 44)
		Me.OptBetter.Name = "OptBetter"
		Me.OptBetter.Size = New System.Drawing.Size(53, 17)
		Me.OptBetter.TabIndex = 2
		Me.OptBetter.TabStop = True
		Me.OptBetter.Text = "Better"
		Me.OptBetter.UseVisualStyleBackColor = True
		'
		'OptFaster
		'
		Me.OptFaster.AutoSize = True
		Me.OptFaster.Location = New System.Drawing.Point(8, 21)
		Me.OptFaster.Name = "OptFaster"
		Me.OptFaster.Size = New System.Drawing.Size(54, 17)
		Me.OptFaster.TabIndex = 1
		Me.OptFaster.TabStop = True
		Me.OptFaster.Text = "Faster"
		Me.OptFaster.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.OptFaster.UseVisualStyleBackColor = True
		'
		'lblPacker
		'
		Me.lblPacker.AutoSize = True
		Me.lblPacker.Location = New System.Drawing.Point(6, 5)
		Me.lblPacker.Name = "lblPacker"
		Me.lblPacker.Size = New System.Drawing.Size(81, 13)
		Me.lblPacker.TabIndex = 0
		Me.lblPacker.Text = "Default Packer:"
		'
		'ChkToolTips
		'
		Me.ChkToolTips.AutoSize = True
		Me.ChkToolTips.Location = New System.Drawing.Point(675, 255)
		Me.ChkToolTips.Name = "ChkToolTips"
		Me.ChkToolTips.Size = New System.Drawing.Size(97, 17)
		Me.ChkToolTips.TabIndex = 124
		Me.ChkToolTips.Text = "Show ToolTips"
		Me.ChkToolTips.UseVisualStyleBackColor = True
		'
		'ChkExpand
		'
		Me.ChkExpand.AutoSize = True
		Me.ChkExpand.Location = New System.Drawing.Point(675, 232)
		Me.ChkExpand.Name = "ChkExpand"
		Me.ChkExpand.Size = New System.Drawing.Size(88, 17)
		Me.ChkExpand.TabIndex = 123
		Me.ChkExpand.Text = "Show Details"
		Me.ChkExpand.UseVisualStyleBackColor = True
		'
		'BtnPartDown
		'
		Me.BtnPartDown.Location = New System.Drawing.Point(675, 174)
		Me.BtnPartDown.Name = "BtnPartDown"
		Me.BtnPartDown.Size = New System.Drawing.Size(96, 23)
		Me.BtnPartDown.TabIndex = 122
		Me.BtnPartDown.Text = "Move Part Down"
		Me.BtnPartDown.UseVisualStyleBackColor = True
		'
		'BtnPartUp
		'
		Me.BtnPartUp.Location = New System.Drawing.Point(675, 145)
		Me.BtnPartUp.Name = "BtnPartUp"
		Me.BtnPartUp.Size = New System.Drawing.Size(96, 23)
		Me.BtnPartUp.TabIndex = 121
		Me.BtnPartUp.Text = "Move Part Up"
		Me.BtnPartUp.UseVisualStyleBackColor = True
		'
		'BtnFileDown
		'
		Me.BtnFileDown.Location = New System.Drawing.Point(675, 173)
		Me.BtnFileDown.Name = "BtnFileDown"
		Me.BtnFileDown.Size = New System.Drawing.Size(96, 23)
		Me.BtnFileDown.TabIndex = 120
		Me.BtnFileDown.Text = "Move File Down"
		Me.BtnFileDown.UseVisualStyleBackColor = True
		Me.BtnFileDown.Visible = False
		'
		'BtnFileUp
		'
		Me.BtnFileUp.Location = New System.Drawing.Point(675, 144)
		Me.BtnFileUp.Name = "BtnFileUp"
		Me.BtnFileUp.Size = New System.Drawing.Size(96, 23)
		Me.BtnFileUp.TabIndex = 119
		Me.BtnFileUp.Text = "Move File Up"
		Me.BtnFileUp.UseVisualStyleBackColor = True
		Me.BtnFileUp.Visible = False
		'
		'txtEdit
		'
		Me.txtEdit.BackColor = System.Drawing.SystemColors.Window
		Me.txtEdit.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.txtEdit.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtEdit.Location = New System.Drawing.Point(675, 379)
		Me.txtEdit.Name = "txtEdit"
		Me.txtEdit.Size = New System.Drawing.Size(28, 16)
		Me.txtEdit.TabIndex = 112
		Me.txtEdit.Visible = False
		'
		'Label10
		'
		Me.Label10.AutoSize = True
		Me.Label10.Location = New System.Drawing.Point(10, 4)
		Me.Label10.Name = "Label10"
		Me.Label10.Size = New System.Drawing.Size(84, 13)
		Me.Label10.TabIndex = 118
		Me.Label10.Text = "Demo Structure:"
		'
		'BtnNew
		'
		Me.BtnNew.Location = New System.Drawing.Point(675, 20)
		Me.BtnNew.Name = "BtnNew"
		Me.BtnNew.Size = New System.Drawing.Size(96, 26)
		Me.BtnNew.TabIndex = 113
		Me.BtnNew.Text = "New Script"
		Me.BtnNew.UseVisualStyleBackColor = True
		'
		'BtnCancel
		'
		Me.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.BtnCancel.Location = New System.Drawing.Point(675, 535)
		Me.BtnCancel.Name = "BtnCancel"
		Me.BtnCancel.Size = New System.Drawing.Size(96, 26)
		Me.BtnCancel.TabIndex = 117
		Me.BtnCancel.Text = "Close"
		Me.BtnCancel.UseVisualStyleBackColor = True
		'
		'BtnOK
		'
		Me.BtnOK.AccessibleRole = System.Windows.Forms.AccessibleRole.None
		Me.BtnOK.Location = New System.Drawing.Point(675, 503)
		Me.BtnOK.Name = "BtnOK"
		Me.BtnOK.Size = New System.Drawing.Size(96, 26)
		Me.BtnOK.TabIndex = 116
		Me.BtnOK.Text = "Close && Build"
		Me.BtnOK.UseVisualStyleBackColor = True
		'
		'BtnSave
		'
		Me.BtnSave.Location = New System.Drawing.Point(675, 84)
		Me.BtnSave.Name = "BtnSave"
		Me.BtnSave.Size = New System.Drawing.Size(96, 26)
		Me.BtnSave.TabIndex = 115
		Me.BtnSave.Text = "Save Script"
		Me.BtnSave.UseVisualStyleBackColor = True
		'
		'BtnLoad
		'
		Me.BtnLoad.Location = New System.Drawing.Point(675, 52)
		Me.BtnLoad.Name = "BtnLoad"
		Me.BtnLoad.Size = New System.Drawing.Size(96, 26)
		Me.BtnLoad.TabIndex = 114
		Me.BtnLoad.Text = "Load Script"
		Me.BtnLoad.UseVisualStyleBackColor = True
		'
		'TV
		'
		Me.TV.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TV.Indent = 19
		Me.TV.Location = New System.Drawing.Point(10, 21)
		Me.TV.Name = "TV"
		Me.TV.Size = New System.Drawing.Size(646, 542)
		Me.TV.TabIndex = 111
		'
		'strip
		'
		Me.strip.ImageScalingSize = New System.Drawing.Size(20, 20)
		Me.strip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tssLabel, Me.TssDisk})
		Me.strip.Location = New System.Drawing.Point(0, 579)
		Me.strip.Name = "strip"
		Me.strip.Size = New System.Drawing.Size(784, 22)
		Me.strip.SizingGrip = False
		Me.strip.TabIndex = 126
		Me.strip.Text = "StatusStrip1"
		'
		'tssLabel
		'
		Me.tssLabel.Name = "tssLabel"
		Me.tssLabel.Size = New System.Drawing.Size(647, 17)
		Me.tssLabel.Spring = True
		Me.tssLabel.Text = "Script:"
		Me.tssLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'TssDisk
		'
		Me.TssDisk.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
		Me.TssDisk.Name = "TssDisk"
		Me.TssDisk.Size = New System.Drawing.Size(122, 17)
		Me.TssDisk.Text = "Disk 1: 664 blocks free"
		'
		'FrmEditor
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(784, 601)
		Me.Controls.Add(Me.strip)
		Me.Controls.Add(Me.PnlPacker)
		Me.Controls.Add(Me.ChkToolTips)
		Me.Controls.Add(Me.ChkExpand)
		Me.Controls.Add(Me.BtnPartDown)
		Me.Controls.Add(Me.BtnPartUp)
		Me.Controls.Add(Me.BtnFileDown)
		Me.Controls.Add(Me.BtnFileUp)
		Me.Controls.Add(Me.txtEdit)
		Me.Controls.Add(Me.Label10)
		Me.Controls.Add(Me.BtnNew)
		Me.Controls.Add(Me.BtnCancel)
		Me.Controls.Add(Me.BtnOK)
		Me.Controls.Add(Me.BtnSave)
		Me.Controls.Add(Me.BtnLoad)
		Me.Controls.Add(Me.TV)
		Me.MinimizeBox = False
		Me.MinimumSize = New System.Drawing.Size(800, 640)
		Me.Name = "FrmEditor"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "FrmEditor"
		Me.PnlPacker.ResumeLayout(False)
		Me.PnlPacker.PerformLayout()
		Me.strip.ResumeLayout(False)
		Me.strip.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents PnlPacker As Panel
	Friend WithEvents OptBetter As RadioButton
	Friend WithEvents OptFaster As RadioButton
	Friend WithEvents lblPacker As Label
	Friend WithEvents ChkToolTips As CheckBox
	Friend WithEvents ChkExpand As CheckBox
	Friend WithEvents BtnPartDown As Button
	Friend WithEvents BtnPartUp As Button
	Friend WithEvents BtnFileDown As Button
	Friend WithEvents BtnFileUp As Button
	Friend WithEvents txtEdit As TextBox
	Friend WithEvents Label10 As Label
	Friend WithEvents BtnNew As Button
	Friend WithEvents BtnCancel As Button
	Friend WithEvents BtnOK As Button
	Friend WithEvents BtnSave As Button
	Friend WithEvents BtnLoad As Button
	Friend WithEvents TV As TreeView
	Friend WithEvents strip As StatusStrip
	Friend WithEvents tssLabel As ToolStripStatusLabel
	Friend WithEvents TssDisk As ToolStripStatusLabel
End Class
