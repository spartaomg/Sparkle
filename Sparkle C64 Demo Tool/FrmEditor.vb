﻿Imports System.ComponentModel

Public Class FrmEditor

	Private ReadOnly colDisk As Color = Color.DarkRed
	Private ReadOnly colDiskGray As Color = Color.FromArgb(41, 41, 41)
	Private ReadOnly colDiskInfo As Color = Color.DarkGreen
	Private ReadOnly colDiskInfoGray As Color = Color.FromArgb(29, 29, 29)
	Private ReadOnly colBunlde As Color = Color.DarkMagenta
	Private ReadOnly colBundleGray As Color = Color.FromArgb(49, 49, 49)
	Private ReadOnly colFile As Color = Color.Black
	Private ReadOnly colFileGray As Color = Color.FromArgb(20, 20, 20)
	Private ReadOnly colFileParamDefault As Color = Color.RosyBrown
	Private ReadOnly colFileParamDefaultGray As Color = Color.FromArgb(96, 96, 96)
	Private ReadOnly colFileParamEdited As Color = Color.SaddleBrown
	Private ReadOnly colFileParamEditedGray As Color = Color.FromArgb(32, 32, 32)
	Private ReadOnly colFileIODefault As Color = Color.MediumPurple
	Private ReadOnly colFileIODefaultGray As Color = Color.FromArgb(96, 96, 96)
	Private ReadOnly colFileIOEdited As Color = Color.Purple
	Private ReadOnly colFileIOEditedGray As Color = Color.FromArgb(32, 32, 32)
	Private ReadOnly colFileSize As Color = Color.DarkGray
	Private ReadOnly colFileSizeGray As Color = Color.FromArgb(67, 67, 67)
	Private ReadOnly colNewBlockNo As Color = Color.MediumPurple
	Private ReadOnly colNewBlockNoGray As Color = Color.FromArgb(96, 96, 96)
	Private ReadOnly colNewBlockYes As Color = Color.Purple
	Private ReadOnly colNewBlockYesGray As Color = Color.FromArgb(32, 32, 32)
	Private ReadOnly colScript As Color = Color.DarkOliveGreen
	Private ReadOnly colScriptGray As Color = Color.FromArgb(35, 35, 35)
	Private ReadOnly colNewEntry As Color = Color.Navy

	Private WithEvents SCC As SubClassCtrl.SubClassing
	Private Const WM_VSCROLL As Integer = &H115
	Private Const WM_HSCROLL As Integer = &H114
	Private Const WM_MOUSEWHEEL As Integer = &H20A

	Private txtBuffer As String = ""
	Private Loading As Boolean = False

	Private DFA, DFO, DFL As Boolean
	Private DFAS, DFOS, DFLS As String
	Private DFAN, DFON, DFLN As Integer
	Private FAS, FOS, FLS As String

	Private Ext As String   'File extension
	Private P() As Byte     'File

	Private FAddr, FOffs, FLen, PLen As Integer

	Private SC, DC, PC, FC As Integer

	Private FirstBundleOfDisk As Boolean

	Private FileType As Byte

	Private LMD As Date = Date.Now
	Private Dbl As Boolean = False

	Private NewD As Boolean = False
	Private HandleKey As Boolean = False
	Private NodeType As Byte
	Private FileSize As Integer
	Private FilePath As String

	Private ReadOnly BaseScriptText As String = "This Script: "
	Private ReadOnly BaseScriptKey As String = "BaseScript"

	Private ReadOnly NewEntryText As String = "Add..."
	Private ReadOnly NewDiskEntryText As String = "New Disk"
	Private ReadOnly NewBundleEntryText As String = "New Bundle"
	Private ReadOnly NewScriptEntryText As String = "New Script"
	Private ReadOnly NewFileEntryText As String = "Add New File"
	Private ReadOnly NewEntryKey As String = "NewEntry"
	Private ReadOnly NewDiskKey As String = "NewDisk"
	Private ReadOnly NewBundleKey As String = "NewBundle"
	Private ReadOnly NewScriptKey As String = "NewScript"
	Private ReadOnly NewFileKey As String = "NewFile"

	Private ZPSet As Boolean = False
	Private LoopSet As Boolean = False

	Private BaseNode, SelNode, NewEntryNode As TreeNode
	Private LoopNode As New TreeNode
	Private ZPNode As New TreeNode
	Private DiskNode As TreeNode
	Private BundleNode As TreeNode
	Private FileNode As TreeNode
	Private ScriptNode As TreeNode
	Private DiskNodeA() As TreeNode

	Private BundleInNewBlock As Boolean = False

	Private ReadOnly sDiskPath As String = "Disk Path: "
	Private ReadOnly sDiskHeader As String = "Disk Header: "
	Private ReadOnly sDiskID As String = "Disk ID: "
	Private ReadOnly sDemoName As String = "Demo Name: "
	Private ReadOnly sDemoStart As String = "Demo Start: "
	Private ReadOnly sAddEntry As String = "AddEntry"
	Private ReadOnly sAddScript As String = "AddScript"
	Private ReadOnly sAddDisk As String = "AddDisk"
	Private ReadOnly sAddBundle As String = "AddBundle"
	Private ReadOnly sAddFile As String = "AddFile"
	Private ReadOnly sFileSize As String = "Original File Size: "
	Private ReadOnly sFileUIO As String = "Load to:      "
	Private ReadOnly sFileAddr As String = "Load Address: $"
	Private ReadOnly sFileOffs As String = "File Offset:  $"
	Private ReadOnly sFileLen As String = "File Length:  $"
	Private ReadOnly sDirArt As String = "DirArt: "
	Private ReadOnly sZP As String = "Zeropage: "
	Private ReadOnly sLoop As String = "Loop: "

	Private ReadOnly sIL0 As String = "Interleave 0: "
	Private ReadOnly sIL1 As String = "Interleave 1: "
	Private ReadOnly sIL2 As String = "Interleave 2: "
	Private ReadOnly sIL3 As String = "Interleave 3: "

	Private ReadOnly sScript As String = "Script: "
	Private ReadOnly sFile As String = "File: "
	Private ReadOnly sNewBlock As String = "Start the first file of this bundle in a new sector on the disk: " '"Start bundle in a new block on the disk: "
	Private ReadOnly TT As New ToolTip

	Private ReadOnly tDiskPath As String = "Double click or press <Enter> to specify where your demo disk will be saved in D64 format."
	Private ReadOnly tDiskHeader As String = "Double click or start typing to edit the disk's header."
	Private ReadOnly tDiskID As String = "Double click or start typing to edit the disk's ID."
	Private ReadOnly tDemoName As String = "Double click or start typing to edit the name of the first PRG in the directory."
	Private ReadOnly tDemoStart As String = "Double click orstart typing to edit the start address (entry point) of the demo."
	Private ReadOnly tAddDisk As String = "Double click or press <Enter> to add a new disk structure to this script."
	Private ReadOnly tAddBundle As String = "Double click or press <Enter> to add a new bundle to this script."
	Private ReadOnly tAddScript As String = "Double click or press <Enter> to add an existing script to this script."
	Private ReadOnly tAddFile As String = "Double click or press <Enter> to add an existing file to this bundle."
	Private ReadOnly tNB As String = "Double click or press <Enter> to change the alignment of this bundle with a sector on the disk."
	Private ReadOnly tFileSize As String = "Original size of the selected file segment in blocks."
	Private ReadOnly tFileAddr As String = "Double click or start typing to edit the file segment's load address."
	Private ReadOnly tFileOffs As String = "Double click or start typing to edit the file segment's offset."
	Private ReadOnly tFileLen As String = "Double click or start typing to edit the file segment's length."
	Private ReadOnly tLoadUIO As String = "Double click or press <Enter> to change the file segment's I/O status, if applicable."
	Private ReadOnly tDirArt As String = "Double click or press <Enter> to add a DirArt file to the demo's directory." + vbNewLine +
				"Press the <Delete> key to delete the current DirArt file."
	Private ReadOnly tDisk As String = "Press <Delete> to delete this disk with all its content."
	Private ReadOnly tBundle As String = "Files and file segments in this bundle will be loaded during a single loader call." + vbNewLine +
				"Press <Delete> to delete this bundle of files from this disk with all its content."
	Private ReadOnly tFile As String = "Double click or press <Enter> to change this file." + vbNewLine +
				"Press <Delete> to delete this file from this bundle."
	Private ReadOnly tScript As String = "Double click or press <Enter> to change this embedded script." + vbNewLine +
				"Press <Delete> to delete this script entry and all its content."
	Private ReadOnly tBaseScript As String = "Double click or press <Enter> to load a script file."
	Private ReadOnly tZP As String = "Double click or press <Enter> to edit the loader's zeropage usage."
	Private ReadOnly tLoop As String = "Double click or press <Enter> to specify the disk the demo will loop to after finishing the last disk." + vbNewLine +
				"The default value of 0 will terminate the demo without looping. A value between 1-255 will result in looping to the specified disk" + vbNewLine +
				"E.g. select 1 to loop to the first disk."

	Private Sub FrmEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		On Error GoTo Err

		'---------------------------------------------------------------
		'       SETTINGS
		'---------------------------------------------------------------

		If My.Settings.EditorWindowMax = True Then
			WindowState = FormWindowState.Maximized
		Else
			WindowState = FormWindowState.Normal
			Width = My.Settings.EditorWidth
			Height = My.Settings.EditorHeight

			Dim R As Rectangle = Screen.FromPoint(Location).WorkingArea

			Dim X = R.Left + (R.Width - Width) \ 2
			Dim Y = R.Top + (R.Height - Height) \ 2
			Location = New Point(X, Y)
		End If

		'If My.Settings.DefaultPacker = 1 Then
		'OptFaster.Checked = True
		'Else
		'OptBetter.Checked = True
		'End If

		If My.Settings.FilePaths = 1 Then
			OptRelativePaths.Checked = True
		Else
			OptFullPaths.Checked = True
		End If

		ChkExpand.Checked = My.Settings.ShowFileDetails
		ChkToolTips.Checked = My.Settings.ShowToolTips
		ChkSize.Checked = My.Settings.CalculateSize

		'---------------------------------------------------------------

		Refresh()

		LoopSet = False
		ZPSet = False

		AddBaseNode()

		SC = 0 : DC = 0 : PC = 0 : FC = 0

		NewD = False

		If Script <> "" Then
			If ScriptName <> "" Then tssLabel.Text = "Script: " + ScriptName
			ConvertScriptToNodes()
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmEditor_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
		On Error GoTo Err

		Dim DropFiles() As String = e.Data.GetData(DataFormats.FileDrop)
		For Each Path In DropFiles
			Select Case Strings.Right(Path, 4)
				Case ".sls"
					SetScriptPath(Path)

					OpenScript()
					TV_GotFocus(sender, e)
			End Select
		Next

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmEditor_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
		On Error GoTo Err

		If e.Data.GetDataPresent(DataFormats.FileDrop) Then
			e.Effect = DragDropEffects.Copy
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmEditor_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
		On Error GoTo Err

		If txtEdit.Visible Then TV.Focus()

		With My.Settings
			'.DefaultPacker = If(OptFaster.Checked, 1, 2)
			.ShowFileDetails = ChkExpand.Checked
			.ShowToolTips = ChkToolTips.Checked
			.CalculateSize = ChkSize.Checked
			.EditorWindowMax = Me.WindowState = FormWindowState.Maximized
			If WindowState = FormWindowState.Normal Then
				.EditorHeight = Height
				.EditorWidth = Width
			End If
		End With

		'----------------------
		ConvertNodesToScript()
		'----------------------

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmEditor_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
		On Error GoTo Err

		If e.Control Then
			Select Case e.KeyCode
				Case Keys.N
					e.SuppressKeyPress = True
					BtnNew_Click(sender, e)
				Case Keys.L
					e.SuppressKeyPress = True
					BtnLoad_Click(sender, e)
				Case Keys.S
					e.SuppressKeyPress = True
					BtnSave_Click(sender, e)
				Case Keys.R
					e.SuppressKeyPress = True
					OptRelativePaths.Checked = True
				Case Keys.P
					e.SuppressKeyPress = True
					OptFullPaths.Checked = True
				Case Keys.PageUp
					e.SuppressKeyPress = True
					BtnEntryUp_Click(sender, e)
				Case Keys.PageDown
					e.SuppressKeyPress = True
					BtnEntryDown_Click(sender, e)
				Case Keys.D
					e.SuppressKeyPress = True
					ChkExpand.Checked = Not ChkExpand.Checked
				Case Keys.T
					e.SuppressKeyPress = True
					ChkToolTips.Checked = Not ChkToolTips.Checked
				Case Keys.A
					e.SuppressKeyPress = True
					ChkSize.Checked = Not ChkSize.Checked
			End Select
		ElseIf e.Shift Then
		ElseIf e.Alt Then
			Select Case e.KeyCode
				Case Keys.F4
					e.SuppressKeyPress = True
					BtnCancel_Click(sender, e)
			End Select
		Else
			Select Case e.KeyCode
				Case Keys.F5
					e.SuppressKeyPress = True
					BtnOK_Click(sender, e)
				Case Else
					Exit Select
			End Select
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FrmEditor_Resize(sender As Object, e As EventArgs) Handles Me.Resize
		On Error GoTo Err

		With TV
			.Width = Width - BtnNew.Width - 56
			.Height = Height - .Top - strip.Height - 56
		End With

		With BtnNew
			.Left = Width - .Width - 32
			BtnLoad.Left = .Left
			BtnSave.Left = .Left
			BtnEntryUp.Left = .Left
			BtnEntryDown.Left = .Left
			ChkExpand.Left = .Left
			ChkToolTips.Left = .Left
			ChkSize.Left = .Left
			'PnlPacker.Left = .Left - 11
			PnlPath.Left = .Left - 11
		End With

		With BtnCancel
			.Top = TV.Top + TV.Height - .Height
			.Left = BtnNew.Left
		End With

		With BtnOK
			.Top = BtnCancel.Top - .Height - 6
			.Left = BtnNew.Left
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TV_KeyDown(sender As Object, e As KeyEventArgs) Handles TV.KeyDown
		On Error GoTo Err

		Dim S As String
		Dim N As TreeNode = TV.SelectedNode

		If e.KeyCode = Keys.Delete Then
			HandleKey = True
			e.SuppressKeyPress = True
			DeleteNode(SelNode)
			Exit Sub
		End If

		SelNode = TV.SelectedNode

		If e.KeyCode = Keys.Enter Then
			Select Case SelNode.ForeColor
				Case colDiskGray, colDiskInfoGray
					HandleKey = True
					e.SuppressKeyPress = True
					Exit Sub
				Case colFileGray, colFileParamDefaultGray, colFileParamEditedGray, colFileIODefaultGray, colFileIOEditedGray, colFileSizeGray
					'File and File Parameters in script
					HandleKey = True
					e.SuppressKeyPress = True
					Exit Sub
				Case colBundleGray, colNewBlockYesGray, colNewBlockNoGray, colScriptGray
					'Disk, Disk info, Bundle, New Block, and Script in script
					HandleKey = True
					e.SuppressKeyPress = True
					Exit Sub
			End Select
		End If

		Select Case SelNode.Text
			Case NewDiskEntryText
				If e.KeyCode = Keys.Enter Then
					HandleKey = True
					e.SuppressKeyPress = True
					AddDiskNode()
				Else
					HandleKey = False
				End If
				Exit Sub
			Case NewBundleEntryText
				If e.KeyCode = Keys.Enter Then
					HandleKey = True
					e.SuppressKeyPress = True
					AddBundleNode()
				Else
					HandleKey = False
				End If
				Exit Sub
			Case NewScriptEntryText
				If e.KeyCode = Keys.Enter Then
					HandleKey = True
					e.SuppressKeyPress = True
					AddScriptNode()
				Else
					HandleKey = False
				End If
				Exit Sub
			Case NewFileEntryText
				If e.KeyCode = Keys.Enter Then
					HandleKey = True
					e.SuppressKeyPress = True
					AddFileNode()
				Else
					HandleKey = False
				End If
				Exit Sub
		End Select

		If SelNode.Name = BaseScriptKey Then
			If e.KeyCode = Keys.Enter Then
				HandleKey = True
				e.SuppressKeyPress = True
				BtnLoad_Click(sender, e)
			Else
				HandleKey = False
			End If
			Exit Sub
		End If

		S = Strings.Left(N.Name, InStr(N.Name, ":") + 1)

		If Strings.Right(N.Name, 3) = ":FS" Then        'Is this a File Size node?
			HandleKey = False
			Exit Sub
		ElseIf Strings.Right(N.Name, 3) = ":NB" Then
			If e.KeyCode = Keys.Enter Then
				HandleKey = True
				e.SuppressKeyPress = True
				txtEdit.Visible = False
				SwapNewBlockStatus()
			Else
				HandleKey = False
			End If
			Exit Sub
		ElseIf Strings.Right(N.Name, 5) = ":FUIO" Then
			If e.KeyCode = Keys.Enter Then
				HandleKey = True
				e.SuppressKeyPress = True
				txtEdit.Visible = False
				SwapIOStatus()
			Else
				HandleKey = False
			End If
			Exit Sub
		ElseIf Strings.Right(N.Name, 3) = ":FA" Then

			'Load Address
			Select Case e.KeyCode
				Case Keys.D0 To Keys.D9, Keys.A To Keys.F, Keys.Enter
					HandleKey = True
					e.SuppressKeyPress = True
					S = sFileAddr
FileData:
					txtEdit.MaxLength = 4
FileDataFO:
					With txtEdit
						.Tag = S
						.Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
						N.Text = S
						.Top = TV.Top + N.Bounds.Top + 3
						.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
						.Width = TextRenderer.MeasureText(.Text, N.NodeFont).Width
						.ForeColor = N.ForeColor
						.Visible = True
					End With
				Case Else
					HandleKey = False
					Exit Sub
			End Select
		ElseIf Strings.Right(N.Name, 3) = ":FO" Then

			'File Offset
			Select Case e.KeyCode
				Case Keys.D0 To Keys.D9, Keys.A To Keys.F, Keys.Enter
					HandleKey = True
					e.SuppressKeyPress = True
					S = sFileOffs
					txtEdit.MaxLength = 8
					GoTo FileDataFO
				Case Else
					HandleKey = False
					Exit Sub
			End Select

		ElseIf Strings.Right(N.Name, 3) = ":FL" Then

			'File Length
			Select Case e.KeyCode
				Case Keys.D0 To Keys.D9, Keys.A To Keys.F, Keys.Enter
					HandleKey = True
					e.SuppressKeyPress = True
					S = sFileLen
					GoTo FileData
				Case Else
					HandleKey = False
					Exit Sub
			End Select

		ElseIf Strings.InStr(N.Name, ":F") > 0 Then   'Is this a File node?

			Select Case e.KeyCode
				Case Keys.Enter
					HandleKey = True
					e.SuppressKeyPress = True

					txtEdit.Visible = False
					FilePath = N.Text
					UpdateFileNode()

				Case Else
					HandleKey = False
			End Select
			Exit Sub

		ElseIf Strings.Left(LCase(N.Text), 8) = "[script:" Then

			Select Case e.KeyCode
				Case Keys.Enter
					HandleKey = True
					e.SuppressKeyPress = True

					txtEdit.Visible = False
					S = sScript
					FilePath = Strings.Right(N.Text.TrimStart("[").TrimEnd("]"), Len(N.Text) - Len(S) - 2)
					UpdateScriptNode(N)

				Case Else
					HandleKey = False
			End Select
			Exit Sub

		Else
			'Any other nodes
			With txtEdit
				.MaxLength = 32767
				.Visible = False    'True
				Select Case S
					Case sDiskPath

						Select Case e.KeyCode
							Case Keys.Enter

								FilePath = Strings.Right(N.Text, Len(N.Text) - Len(S))
								.Tag = S
								HandleKey = True
								e.SuppressKeyPress = True
								UpdateDiskPath()
							Case Else
								HandleKey = False
						End Select

						Exit Sub

					Case sDiskHeader

						Select Case e.KeyCode

							Case Keys.A To Keys.Z, Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.Enter
								.Tag = S
								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
								N.Text = S
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.Width = TV.Left + TV.Width - .Left - 2 - 17    'Subtract border and scrollbar widths
								.MaxLength = 16
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select

					Case sDiskID

						Select Case e.KeyCode
							Case Keys.A To Keys.Z, Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.Enter

								.Tag = S
								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
								N.Text = S
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.Width = TV.Left + TV.Width - .Left - 2 - 17    'Subtract border and scrollbar widths
								.MaxLength = 5
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select

					Case sDemoName

						Select Case e.KeyCode
							Case Keys.A To Keys.Z, Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.Enter

								.Tag = S
								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
								N.Text = S
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.Width = TV.Left + TV.Width - .Left - 2 - 17    'Subtract border and scrollbar widths
								.MaxLength = 16
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select

					Case sDemoStart

						Select Case e.KeyCode
							Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.A To Keys.F, Keys.Enter

								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S) - 1)
								.Width = TextRenderer.MeasureText("0000", N.NodeFont).Width
								.Tag = S + "$"
								N.Text = .Tag
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.MaxLength = 4
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select

					Case sDirArt
						Select Case e.KeyCode
							Case Keys.Enter

								FilePath = Strings.Right(N.Text, Len(N.Text) - Len(S))
								.Tag = S
								.Visible = False
								HandleKey = True
								e.SuppressKeyPress = True

								UpdateDirArtPath()
							Case Else
								HandleKey = False
						End Select

						Exit Sub

					Case sZP

						Select Case e.KeyCode
							Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.A To Keys.F, Keys.Enter

								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S) - 1)
								.Width = TextRenderer.MeasureText("00", N.NodeFont).Width
								.Tag = S + "$"
								N.Text = .Tag
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.MaxLength = 2
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select

						'Case sPacker
						'Select Case e.KeyCode
						'Case Keys.Enter

						'.Visible = False
						'Select Case LCase(Strings.Right(N.Text, 6))
						'Case "faster"
						'N.Text = sPacker + "better"
						'Packer = 2
						'Case "better"
						'N.Text = sPacker + "faster"
						'Packer = 1
						'End Select

						'HandleKey = True
						'e.SuppressKeyPress = True
						'----------------
						'CalcDiskSizeWithForm(BaseNode, -1) ' SelNode.Parent.Index + 1)
						''----------------
						'Case Else
						'HandleKey = False
						'End Select

						'Exit Sub

					Case sLoop

						Select Case e.KeyCode
							Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.Enter
								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
								.Width = TextRenderer.MeasureText("000", N.NodeFont).Width
								.Tag = S
								N.Text = .Tag
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.MaxLength = 3
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select

					Case sIL0, sIL1, sIL2, sIL3
						'If CustomIL Then
						Select Case e.KeyCode
							Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9, Keys.Enter
								.Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
								.Width = TextRenderer.MeasureText("00", N.NodeFont).Width
								.Tag = S
								N.Text = .Tag
								.Left = TV.Left + N.Bounds.Left + N.Bounds.Width
								.MaxLength = 2
								.Visible = True

								HandleKey = True
								e.SuppressKeyPress = True
							Case Else
								HandleKey = False
								Exit Sub
						End Select
						'end If
					Case Else
						HandleKey = False
						.Visible = False
						Exit Sub
				End Select
				.Top = TV.Top + N.Bounds.Top + 3
				.ForeColor = N.ForeColor
			End With
		End If
		With txtEdit
			If txtEdit.Visible = True Then
				txtBuffer = .Text
				'txtEdit.Text = Char.ConvertFromUtf32(e.KeyValue)
				If e.KeyCode = Keys.Enter Then
					.SelectAll()
				Else
					If e.Modifiers = Keys.Shift Then
						.Text = ChrW(e.KeyValue)
					Else
						.Text = LCase(ChrW(e.KeyValue))
					End If
					.SelectionStart = 1 'Len(txtEdit.Text)
					.SelectionLength = 0
				End If
				'.Refresh()
				.Focus()
			End If
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TV.KeyPress
		On Error GoTo Err

		If HandleKey = True Then e.Handled = True

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TV_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TV.AfterSelect
		On Error GoTo Err

		NodeSelect()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
	End Sub

	Private Sub TV_GotFocus(sender As Object, e As EventArgs) Handles TV.GotFocus
		On Error GoTo Err

		NodeSelect()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TV.NodeMouseClick
		On Error GoTo Err

		If e.Button = MouseButtons.Right Then
			TV.SelectedNode = e.Node
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TV.NodeMouseDoubleClick
		On Error GoTo Err

		Dim k As New KeyEventArgs(Keys.Enter)
		TV_KeyDown(sender, k)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_MouseDown(sender As Object, e As MouseEventArgs) Handles TV.MouseDown
		On Error GoTo Err

		'Handle double clicks before a node is expanded or collapsed to prevent unwanted expansion or collapse

		Dim Delta As TimeSpan = Date.Now - LMD

		If Delta.TotalMilliseconds < SystemInformation.DoubleClickTime Then
			Dbl = True
		Else
			Dbl = False
		End If

		LMD = Date.Now

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TV_MouseEnter(sender As Object, e As EventArgs) Handles TV.MouseEnter
		On Error GoTo Err

		With TT
			.ToolTipIcon = ToolTipIcon.Info
			.UseFading = True
			.InitialDelay = 1000
			'.AutomaticDelay = 2000
			.AutoPopDelay = 5000
			.ReshowDelay = 1000
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TV_MouseLeave(sender As Object, e As EventArgs) Handles TV.MouseLeave
		On Error GoTo Err

		TT.Hide(TV)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_DragDrop(sender As Object, e As DragEventArgs) Handles TV.DragDrop
		On Error GoTo Err

		FrmEditor_DragDrop(sender, e)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_DragEnter(sender As Object, e As DragEventArgs) Handles TV.DragEnter
		On Error GoTo Err

		FrmEditor_DragEnter(sender, e)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles TV.BeforeExpand
		On Error GoTo Err

		'If (TV.SelectedNode.Tag < &H10000) And (TV.SelectedNode.Tag > 0) Then
		''If ChkExpand.Checked = False Then
		If Dbl = True Then e.Cancel = True
		''End If
		'End If

		Dbl = False

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_BeforeCollapse(sender As Object, e As TreeViewCancelEventArgs) Handles TV.BeforeCollapse
		On Error GoTo Err

		'If (TV.SelectedNode.Tag < &H10000) And (TV.SelectedNode.Tag > 0) Then
		''If ChkExpand.Checked = True Then
		If Dbl = True Then e.Cancel = True
		''End If
		'End If

		Dbl = False

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TxtEdit_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEdit.KeyDown
		On Error GoTo Err

		Select Case e.KeyCode
			Case Keys.Left, Keys.Right, Keys.Back, Keys.Delete, Keys.End' Keys.Home, Keys.End
			Case Keys.Enter, Keys.Up, Keys.Down, Keys.Home ', Keys.PageUp, Keys.PageDown
				e.SuppressKeyPress = True
				e.Handled = True
				TV.Focus()
				If e.KeyCode = Keys.Up Then
					If SelNode.Index = 0 Then
						TV.SelectedNode = SelNode.Parent
					Else
						TV.SelectedNode = SelNode.Parent.Nodes(SelNode.Index - 1)
					End If
				ElseIf (e.KeyCode = Keys.Down) Or (e.KeyCode = Keys.Enter) Then
					If SelNode.Index = SelNode.Parent.Nodes.Count - 1 Then
						TV.SelectedNode = SelNode.Parent.Parent.Nodes(SelNode.Parent.Index + 1)
					Else
						TV.SelectedNode = SelNode.Parent.Nodes(SelNode.Index + 1)
					End If
				Else
					TV.SelectedNode = TV.TopNode
				End If
			Case Keys.Escape
				txtEdit.Text = txtBuffer
				e.SuppressKeyPress = True
				e.Handled = True
				TV.Focus()
			Case Else
				If (txtEdit.MaxLength = 4) Or (txtEdit.MaxLength = 2) Or (txtEdit.MaxLength = 8) Then   'Hex numbers
					If Strings.Left(txtEdit.Tag, 2) = Strings.Left(sIL0, 2) Then
						'If CustomIL Then
						GoTo Numeric
						'End If
					Else
						Select Case e.KeyCode
							Case Keys.A To Keys.F, Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9
								Select Case e.KeyCode
									Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9
										If (e.Modifiers = Keys.Shift) OrElse (e.Modifiers = Keys.Control) OrElse (e.Modifiers = Keys.Alt) Then
											e.SuppressKeyPress = True
											e.Handled = True
										End If
								End Select
							Case Else
								e.SuppressKeyPress = True
								e.Handled = True
						End Select
					End If
				ElseIf txtEdit.MaxLength = 3 Then   'Loop feature, on decimals
Numeric:            Select Case e.KeyCode
						Case Keys.D0 To Keys.D9, Keys.NumPad0 To Keys.NumPad9
							If (e.Modifiers = Keys.Shift) OrElse (e.Modifiers = Keys.Control) OrElse (e.Modifiers = Keys.Alt) Then
								e.SuppressKeyPress = True
								e.Handled = True
							End If
						Case Else
							e.SuppressKeyPress = True
							e.Handled = True
					End Select
				End If
		End Select

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TxtEdit_GotFocus(sender As Object, e As EventArgs) Handles txtEdit.GotFocus
		On Error GoTo Err

		SCC = New SubClassCtrl.SubClassing(TV.Handle) With {
		.SubClass = True
		}

		'txtBuffer = txtEdit.Text

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TxtEdit_LostFocus(sender As Object, e As EventArgs) Handles txtEdit.LostFocus
		On Error GoTo Err

		SCC.ReleaseHandle()

		TV.Focus()

		CorrectTextLength()

		SelNode.Text = txtEdit.Tag + txtEdit.Text

		TV.Invalidate(SelNode.Bounds)   'This is to repaint selnode

		txtEdit.Visible = False

		If txtEdit.Text = txtBuffer Then Exit Sub

		Select Case Strings.Right(SelNode.Name, 3)
			Case ":FA", ":FO", ":FL"
				ChangeFileParameters(SelNode.Parent, SelNode.Index)
				'CalcDiskSizeWithForm(BaseNode, SelNode.Parent.Parent.Index) 'NOT NEEDED HERE, Called from ValidateFileParams 
				Exit Sub
		End Select

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TxtEdit_MouseWheel(sender As Object, e As MouseEventArgs) Handles txtEdit.MouseWheel
		On Error GoTo Err

		TV.Focus()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub SCC_CallBackProc(ByRef m As Message) Handles SCC.CallBackProc
		On Error GoTo Err

		'If txtEdit is visible while we are scrolling - set focus back to Tv and hide txtEdit
		Select Case m.Msg
			Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
				TV.Focus()
		End Select

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddBaseNode()
		On Error GoTo Err

		StartUpdate()
		BaseNode = TV.Nodes.Add(BaseScriptKey, BaseScriptText + "(New Script)")
		TV.SelectedNode = BaseNode
		With BaseNode
			.Tag = &H30000
			.BackColor = Color.LightGray
		End With
		AddNewEntryNode()

		With ZPNode
			.Name = sZP
			.Text = sZP + "$02"
			.ForeColor = colDiskInfo
			.NodeFont = New Font("Consolas", 10)
		End With

		With LoopNode
			.Name = sLoop
			.Text = sLoop + "0"
			.ForeColor = colDiskInfo
			.NodeFont = New Font("Consolas", 10)
		End With

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		FinishUpdate()
		BaseNode.ExpandAll()

	End Sub

	Private Sub AddNewEntryNode()
		On Error GoTo Err

		NewEntryNode = BaseNode.Nodes.Add(NewEntryKey, NewEntryText)

		With NewEntryNode
			.Tag = 0
			.ForeColor = colNewEntry
			.Nodes.Add(NewDiskKey, NewDiskEntryText)
			.Nodes.Add(NewBundleKey, NewBundleEntryText)
			.Nodes.Add(NewScriptKey, NewScriptEntryText)
			.Nodes(NewDiskKey).ForeColor = colNewEntry
			.Nodes(NewBundleKey).ForeColor = colNewEntry
			.Nodes(NewScriptKey).ForeColor = colNewEntry
			.Nodes(NewDiskKey).Tag = 0
			.Nodes(NewBundleKey).Tag = 0
			.Nodes(NewScriptKey).Tag = 0
			.Expand()
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddNewFileEntryNode(BundleNode As TreeNode)
		On Error GoTo Err

		With BundleNode
			.Nodes.Add(BundleNode.Name + ":" + NewFileKey, NewFileEntryText)
			.Nodes(BundleNode.Name + ":" + NewFileKey).ForeColor = colNewEntry
			.Nodes(BundleNode.Name + ":" + NewFileKey).Tag = .Tag
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddFileNode()
		On Error GoTo Err

		Dim N As TreeNode = TV.SelectedNode

		FileType = 2  'Prg file
		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		FAddr = -1
		FOffs = -1
		FLen = 0

		StartUpdate()

		FC += 1

		CurrentFile = FC

		If N.Index = 0 Then     'Is this the first file in this bundle?
			BlockCnt = 0        'Yes, reset block count
		Else
			BlockCnt = 1        'Fake block for compression
		End If

		With N
			.Text = NewFile
			.Name = .Parent.Name + ":F" + FC.ToString
			.Tag = FC '.Parent.Tag + FC
			.ForeColor = colFile
		End With

		FileNode = N

		GetFile(N.Text)
		GetDefaultFileParameters(N)

		FAddr = DFAN
		FOffs = DFON
		FLen = DFLN

		DFA = True
		DFO = True
		DFL = True

		If OverlapsIO() = True Then
			N.Text += "*"
		End If

		N.Text = CalcFileSize(N.Text, DFAN, DFLN)          'Also sets/clears IOBit

		UpdateFileParameters(N)

		AddNewFileEntryNode(N.Parent)

		'---------------------------------------
		CalcDiskSizeWithForm(BaseNode, N.Parent.Index)
		'---------------------------------------

		BtnFileUp.Enabled = True

		'---------------------------------------
		'If CurrentDisk > 0 Then
		'TssDisk.Text = "Disk " + (CurrentDisk).ToString + ": " + (maxdisksize - DiskSizeA(CurrentDisk - 1)).ToString + " block" + if(maxdisksize - DiskSizeA(CurrentDisk - 1) <> 1, "s free", " free")
		'End If
		'---------------------------------------

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		FinishUpdate()

	End Sub

	Private Sub AddScriptNode()
		On Error GoTo Err

		FileType = 4    'Script file
		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		SC += 1
		ScriptNode = NewEntryNode

		StartUpdate()
		Loading = True

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		With ScriptNode
			.Nodes.Clear()
			.Name = "S" + SC.ToString
			.Text = "[Script: " + NewFile + "]"
			.ForeColor = colScript
			.Tag = SC + &H30000
		End With
		AddNewEntryNode()
		BundleInNewBlock = False
		ConvertScriptToScriptNodes(ScriptNode, NewFile)
		'ReNumberEntries(BaseNode)

		CalcDiskSize(BaseNode, ScriptNode.Index)

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		Frm.Close()
		Loading = False
		FinishUpdate()
		ScriptNode.Expand()

	End Sub

	Private Sub UpdateScriptNode(SN As TreeNode)
		On Error GoTo Err

		FileType = 4    'Script file
		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		StartUpdate()
		Loading = True

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		With SN
			.Nodes.Clear()
			.Text = "[Script: " + NewFile + "]"
			.NodeFont = New Font(TV.Font.Name, TV.Font.Size, FontStyle.Regular)
		End With
		ScriptNode = SN
		BundleInNewBlock = False
		ConvertScriptToScriptNodes(SN, NewFile)

		CalcDiskSize(BaseNode, ScriptNode.Index)

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		Frm.Close()

		Loading = False
		FinishUpdate()
		SN.Expand()

	End Sub

	Private Sub OpenDemoFile()
		On Error GoTo Err

		Dim P, F As String

		Dim N As TreeNode = TV.SelectedNode

		P = Application.ExecutablePath
		F = ""

		If FilePath <> "" Then
			For I = Len(FilePath) To 1 Step -1
				If Mid(FilePath, I, 1) = "\" Then
					P = Strings.Left(FilePath, I - 1)                                   'Path
					F = Replace(Strings.Right(FilePath, Len(FilePath) - I), "*", "")    'File name, delete IO status asterisk
					Exit For
				End If
			Next
		End If

		Select Case FileType
			Case 1  'Text file
				OpenFile("Open DirArt File", "All accepted file formats (*.txt; *.d64; *.prg; *.bin)|*.txt; *.d64; *.prg; *.bin|Text Files (*.txt)|*.txt|D64 Files (*.d64)|*.d64|PRG Files (*.prg)|*.prg|Binary Files (*.bin)|*.bin|All Files (*.*)|*.*", P)
			Case 2  'Prg File
				OpenFile("Open C64 Program File", "All Files (*.*)|*.*|PRG, SID, and Binary Files (*.prg; *.sid; *.bin)|*.prg; *.sid; *.bin|PRG Files (*.prg)|*.prg|Binary Files (*.bin)|*.bin|SID Files (*.sid)|*.sid", P)
			Case 3  'D64 file
				SaveFile("Save D64 Disk Image File", "D64 Files (*.d64)|*.d64", P, F, False)
			Case 4  'Script file
				OpenFile("Add a Script File", "SLS Files (*.sls)|*.sls|Text Files (*.txt)|*.txt|All Files (*.*)|*.*", P)
			Case 5
				SaveFile("Save Script File", "SLS Files (*.sls)|*.sls", P, F, False)
		End Select

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub OpenFile(Optional dlgTitle As String = "", Optional dlgFilter As String = "", Optional dlgPath As String = "", Optional dlgFile As String = "")
		On Error GoTo Err

		If dlgTitle = "" Then
			dlgTitle = "Open"
		End If

		If dlgFilter = "" Then
			dlgFilter = "PRG and Binary Files (*.prg; *.sid; *.bin; *.txt)|*.prg; *.sid; *.bin; *.txt|PRG Files (*.prg)|*.prg|SID Files (*.sid)|*.sid|Binary Files (*.bin; *.txt)|*.bin; *txt"
		End If

		Dim OpenDLG As New OpenFileDialog

		With OpenDLG
			.Title = dlgTitle
			.Filter = dlgFilter
			.FileName = dlgFile
			.RestoreDirectory = True
			Dim R As DialogResult = .ShowDialog(Me)

			If R = DialogResult.OK Then
				NewFile = .FileName
				FilePath = NewFile
			ElseIf R = DialogResult.Cancel Then
				NewFile = ""
			End If
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub SaveFile(Optional dlgTitle As String = "", Optional dlgFilter As String = "", Optional dlgPath As String = "", Optional dlgFile As String = "", Optional OW As Boolean = True)
		On Error GoTo Err

		If dlgTitle = "" Then
			dlgTitle = "Save"
		End If

		If dlgFilter = "" Then
			dlgFilter = "Sparkle Loader Script files (*.sls)|*.sls"
		End If

		If dlgPath = "" Then
			dlgPath = UserFolder + "\OneDrive\C64\Coding"
		End If

		If dlgFile = "" Then
			dlgFile = ScriptName
		End If

		Dim SaveDLG As New SaveFileDialog

		With SaveDLG
			.Title = dlgTitle
			.Filter = dlgFilter
			.FileName = dlgFile
			.OverwritePrompt = OW
			.RestoreDirectory = True
			Dim R As DialogResult = SaveDLG.ShowDialog(Me)

			NewFile = If(R = DialogResult.OK, .FileName, "")
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddNewBlockNode(BunldeNode As TreeNode)
		On Error GoTo Err

		Dim Fnt As New Font("Consolas", 10)
		AddNode(BunldeNode, BunldeNode.Name + ":NB", sNewBlock + "NO", BunldeNode.Tag, colNewBlockNo, Fnt)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddBundleNode()
		On Error GoTo Err

		BundleNode = NewEntryNode    'TV.SelectedNode

		PC += 1
		ReDim Preserve BundleBytePtrA(PC), BundleBitPtrA(PC), BundleBitPosA(PC), BundleSizeA(PC), BundleOrigSizeA(PC)

		CurrentBundle = PC

		StartUpdate()
		With BundleNode
			.Nodes.Clear()
			.Text = "[Bundle " + PC.ToString + "]"
			.Name = "P" + PC.ToString
			.Tag = PC + &H20000
			.ForeColor = colBunlde
		End With


		AddNewBlockNode(BundleNode)      'Adds a "New Block: false" node to this bundle node
		AddNewFileEntryNode(BundleNode)  'Adds an "Add New File" node to this bundle node

		AddNewEntryNode()       'Adds an "Add.." node with 3 child nodes to the BaseNode

		FinishUpdate()

		BundleNode.Expand()
		TV.SelectedNode = BundleNode.Nodes(1)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub
	Private Sub AddDiskNode()
		On Error GoTo Err

		DiskNode = NewEntryNode

		DC += 1

		'If DiskSizeA.Count < DC Then
		'ReDim Preserve DiskSizeA(DC - 1)
		'End If

		CurrentDisk = DC

		StartUpdate()
		With DiskNode
			.Nodes.Clear()
			.Text = "[Disk " + DC.ToString + If(ChkSize.Checked, ": 0 blocks used, 664 blocks free]", "]")
			.Name = "D" + DC.ToString
			.ForeColor = colDisk
			.Tag = DC + &H10000
			.NodeFont = New Font(TV.Font, FontStyle.Bold)
		End With

		Dim Fnt As New Font("Consolas", 10)

		AddNode(DiskNode, sDiskPath + DC.ToString, sDiskPath + "demo.d64", DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sDiskHeader + DC.ToString, sDiskHeader + "demo disk " + Year(Now).ToString, DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sDiskID + DC.ToString, sDiskID + "sprkl", DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sDemoName + DC.ToString, sDemoName + "demo", DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sDemoStart + DC.ToString, sDemoStart + "$", DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sDirArt + DC.ToString, sDirArt, DiskNode.Tag, colDiskInfo, Fnt)
		'AddNode(DiskNode, sPacker + DC.ToString, sPacker + If(My.Settings.DefaultPacker = 1, "faster", "better"), DiskNode.Tag, colDiskInfo, Fnt)
		'If CustomIL Then
		AddNode(DiskNode, sIL0 + DC.ToString, sIL0 + Strings.Left("04", 2 - Len(IL0.ToString)) + IL0.ToString, DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sIL1 + DC.ToString, sIL1 + Strings.Left("03", 2 - Len(IL1.ToString)) + IL1.ToString, DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sIL2 + DC.ToString, sIL2 + Strings.Left("03", 2 - Len(IL2.ToString)) + IL2.ToString, DiskNode.Tag, colDiskInfo, Fnt)
		AddNode(DiskNode, sIL3 + DC.ToString, sIL3 + Strings.Left("03", 2 - Len(IL3.ToString)) + IL3.ToString, DiskNode.Tag, colDiskInfo, Fnt)
		'End If
		If ZPSet = False Then
			ZPNode.Tag = DiskNode.Tag
			DiskNode.Nodes.Add(ZPNode)
			DiskNode.Nodes(sZP).ForeColor = colDiskInfo
			ZPSet = True
		End If
		If LoopSet = False Then
			LoopNode.Tag = DiskNode.Tag
			DiskNode.Nodes.Add(LoopNode)
			DiskNode.Nodes(sLoop).ForeColor = colDiskInfo
			LoopSet = True
		End If
		AddNewEntryNode()

		FinishUpdate()
		TV.SelectedNode = DiskNode
		DiskNode.Expand()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddNode(Parent As TreeNode, Name As String, Text As String, Optional Tag As Integer = 0, Optional NodeColor As Color = Nothing, Optional NodeFnt As Font = Nothing)
		On Error GoTo Err

		Parent.Nodes.Add(Name, Text)
		Parent.Nodes(Name).Tag = Tag

		Parent.Nodes(Name).ForeColor = If(NodeColor = Nothing, colFile, NodeColor)

		If NodeFnt IsNot Nothing Then
			Parent.Nodes(Name).NodeFont = NodeFnt
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub DeleteNode(N As TreeNode)
		On Error GoTo Err

		Dim Frm As New FrmDisk

		Dim P As TreeNode

		Select Case NodeType
			Case 0
				'MsgBox("This entry cannot be deleted!", vbInformation + vbOKOnly)
			Case 1  'Disk
				If MsgBox("Are you sure you want to delete this Disk and all its content?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

					Frm.Show(Me)
					StartUpdate()

					For I As Integer = N.Nodes.Count - 1 To 0 Step -1
						If N.Nodes(I).Name = sLoop Then
							LoopNode = N.Nodes(I)
							LoopSet = False
							Exit For
						End If
					Next
					For I As Integer = N.Nodes.Count - 1 To 0 Step -1
						If N.Nodes(I).Name = sZP Then
							ZPNode = N.Nodes(I)
							ZPSet = False
							Exit For
						End If
					Next
					Dim NI As Integer = N.Index
					N.Remove()
					If ZPSet = False Then
						For I As Integer = 0 To BaseNode.Nodes.Count - 1
							If Strings.Left(BaseNode.Nodes(I).Text, 5) = "[Disk" Then
								BaseNode.Nodes(I).Nodes.Add(ZPNode)
								ZPSet = True
								Exit For
							End If
						Next
					End If
					If LoopSet = False Then
						For I As Integer = 0 To BaseNode.Nodes.Count - 1
							If Strings.Left(BaseNode.Nodes(I).Text, 5) = "[Disk" Then
								BaseNode.Nodes(I).Nodes.Add(LoopNode)
								LoopSet = True
								Exit For
							End If
						Next
					End If
					CalcDiskSize(BaseNode, NI)
					GoTo Done
				End If
			Case 2  'Bundle
				If MsgBox("Are you sure you want to delete this Bundle and all its content?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

					Frm.Show(Me)

					StartUpdate()

					Dim NI = N.Index
					N.Remove()

					If NI > 0 Then NI -= 1

					CalcDiskSize(BaseNode, NI)
					GoTo Done
				End If
			Case 3  'File
				If MsgBox("Are you sure you want to delete the following file entry?" + vbNewLine + vbNewLine + N.Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

					Frm.Show(Me)

					P = N.Parent
					N.Remove()

					StartUpdate()
					CalcDiskSize(BaseNode, P.Index)

					GoTo Done

				End If
			Case 4  'Script
				If MsgBox("Are you sure you want to delete the following script entry?" + vbNewLine + vbNewLine + N.Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
					'P = N.Parent
					Frm.Show(Me)

					StartUpdate()
					Dim NI = N.Index

					N.Remove()

					If NI > 0 Then NI -= 1

					If BaseNode.Nodes.Count > NI Then
						CalcDiskSize(BaseNode, NI)
					End If
					GoTo Done
				End If
			Case 5  'DirArt
				If N.Text <> sDirArt Then
					If MsgBox("Are you sure you want to delete the following DirArt file?" + vbNewLine + vbNewLine + Strings.Right(N.Text, Len(N.Text) - Len(sDirArt)), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
						N.Text = sDirArt
					End If
				End If
		End Select

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:
		Frm.Close()
		FinishUpdate()
		NodeSelect()

	End Sub

	Private Sub NodeSelect()
		On Error GoTo Err

		If TV.SelectedNode Is Nothing Then Exit Sub

		SelNode = TV.SelectedNode

		If (TV.Enabled = False) Or (Loading) Then Exit Sub

		'tssLabel.Text = Hex(SelNode.Tag)

		'Identify nodes that can be deleted
		Select Case SelNode.ForeColor
			Case colDisk        'Disk node
				NodeType = 1
			Case colBunlde      'Bundle node
				NodeType = 2
			Case colFile        'File node
				NodeType = 3
			Case colScript      'Script node
				NodeType = 4
			Case Else
				If Strings.Left(SelNode.Text, 8) = "DirArt: " Then
					NodeType = 5    'DirArt node
				Else
					NodeType = 0    'All other nodes
				End If
		End Select

		Select Case NodeType
			Case 1, 2, 4
				BtnEntryUp.Enabled = SelNode.Index > 0
				BtnEntryDown.Enabled = SelNode.Index < BaseNode.Nodes.Count - 2
			Case Else
				BtnEntryDown.Enabled = False
				BtnEntryUp.Enabled = False
		End Select

		FindCurrentDisk(SelNode, True)

		If CurrentDisk > 0 Then
			If DiskSizeA.Count > CurrentDisk Then
				TssDisk.Text = "Disk " + CurrentDisk.ToString + ": " + (MaxDiskSize - DiskSizeA(CurrentDisk)).ToString + " block" + If(MaxDiskSize - DiskSizeA(CurrentDisk) <> 1, "s free", " free")
			Else
				TssDisk.Text = "Disk " + CurrentDisk.ToString + ": " + MaxDiskSize.ToString + " blocks free"
			End If
		Else
			TssDisk.Text = ""
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FindCurrentDisk(N As TreeNode, Optional FirstRun As Boolean = False)
		On Error GoTo Err

		If FirstRun Then
			CurrentDisk = 0
		End If

		If N.Name = BaseScriptKey Then Exit Sub

		Select Case Int(N.Tag / &H10000)
			Case 0
				'File
				FindCurrentDisk(N.Parent)
			Case 1
				'Disk
				CurrentDisk = N.Tag And &HFFFF
				Exit Sub
			Case 2 To 3
				'Bundle and Script Nodes
				For I As Integer = N.Index To 0 Step -1
					If Int(N.Parent.Nodes(I).Tag / &H10000) = 1 Then
						CurrentDisk = N.Parent.Nodes(I).Tag And &HFFFF
						Exit Sub
					End If
				Next
				If N.Parent.Name <> BaseNode.Name Then
					FindCurrentDisk(N.Parent)
				End If
		End Select

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub StartUpdate()
		On Error GoTo Err

		If Loading = True Then Exit Sub

		Cursor = Cursors.WaitCursor

		With TV
			.Enabled = False
			.BeginUpdate()
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub FinishUpdate()
		On Error GoTo Err

		If Loading = True Then Exit Sub

		With TV
			.EndUpdate()
			.Enabled = True
			.Focus()
		End With

		Cursor = Cursors.Default

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub CorrectTextLength()
		On Error GoTo Err

		'Corrects the length of txtEdit to 2 or 4 characters depending on MaxLength
		'Eliminates invalid values

		'Verify IL, Loop, ZP, and address values (cannot be 0, otherwise IL=Max mod IL)
		Select Case txtEdit.Tag
			Case sIL0
				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If

				Dim IL As Integer = Convert.ToInt32(txtEdit.Text, 10) Mod 21
				If IL = 0 Then IL = DefaultIL0
				txtEdit.Text = IL.ToString

				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If
			Case sIL1
				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If

				Dim IL As Integer = Convert.ToInt32(txtEdit.Text, 10) Mod 19
				If IL = 0 Then IL = DefaultIL1
				txtEdit.Text = IL.ToString

				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If
			Case sIL2
				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If

				Dim IL As Integer = Convert.ToInt32(txtEdit.Text, 10) Mod 18
				If IL = 0 Then IL = DefaultIL2
				txtEdit.Text = IL.ToString

				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If
			Case sIL3
				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If

				Dim IL As Integer = Convert.ToInt32(txtEdit.Text, 10) Mod 17
				If IL = 0 Then IL = DefaultIL3
				txtEdit.Text = IL.ToString

				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("00", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If
			Case sZP + "$"
				If Len(txtEdit.Text) < 2 Then
					txtEdit.Text = Strings.Left("02", 2 - Len(txtEdit.Text)) + txtEdit.Text
				End If

				'Invalid values: 00, 01, ff
				If txtEdit.Text = "00" Or txtEdit.Text = "01" Then
					txtEdit.Text = "02"
				ElseIf txtEdit.Text = "ff" Then
					txtEdit.Text = "fe"
				End If
			Case sLoop
				If txtEdit.Text = "" Then txtEdit.Text = "0"
				Dim L As Integer = Convert.ToInt32(txtEdit.Text, 10)
				If L > 255 Then L = 255
				txtEdit.Text = L.ToString
			Case Else
				Select Case txtEdit.MaxLength
					Case 2  'ZP value, see above
					Case 3  'Loop value, see above
					Case 4  'Address and Length values
						If (Len(txtEdit.Text) < 4) And (Len(txtEdit.Text) > 0) Then
							txtEdit.Text = Strings.Left("0000", 4 - Len(txtEdit.Text)) + txtEdit.Text
						End If
					Case 8  'Offset values
						If (Len(txtEdit.Text) < 8) And (Len(txtEdit.Text) > 0) Then
							txtEdit.Text = Strings.Left("00000000", 8 - Len(txtEdit.Text)) + txtEdit.Text
						End If
					Case Else
						Exit Sub
				End Select
		End Select

		txtEdit.Text = LCase(txtEdit.Text)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ChangeFileParameters(FileNode As TreeNode, NodeIndex As Integer)
		On Error GoTo Err

		StartUpdate()

		GetFile(FileNode.Text)

		'If node is edited, then A/O/L=txtEdit, otherwise A/O/L = 4 rightmost chars of node text
		Dim A As String = If(NodeIndex = 0, txtEdit.Text, Strings.Right(FileNode.Nodes(0).Text, Len(FileNode.Nodes(0).Text) - Len(sFileAddr)))
		Dim O As String = If(NodeIndex = 1, txtEdit.Text, Strings.Right(FileNode.Nodes(1).Text, Len(FileNode.Nodes(1).Text) - Len(sFileOffs)))
		Dim L As String = If(NodeIndex = 2, txtEdit.Text, Strings.Right(FileNode.Nodes(2).Text, Len(FileNode.Nodes(2).Text) - Len(sFileLen)))

		If txtEdit.Text = "" Then
			Select Case NodeIndex
				Case 0
					GetDefaultFileParameters(FileNode)
				Case 1
					GetDefaultFileParameters(FileNode, A)
				Case 2
					GetDefaultFileParameters(FileNode, A, O)
			End Select
			ResetFileParameters(FileNode, NodeIndex)
		ElseIf txtEdit.Text <> txtBuffer Then
			GetDefaultFileParameters(FileNode, A)
			Select Case NodeIndex
				Case 0
					'Load Address has changed, check if Offset is default and update it as needed
					If FileNode.Nodes(1).ForeColor = colFileParamDefault Then FileNode.Nodes(1).Text = sFileOffs + DFOS
					GoTo Node2
				Case 1
					'Load Address and/or Offset have changed, check if length is default and update it as needed
Node2:              If FileNode.Nodes(2).ForeColor = colFileParamDefault Then FileNode.Nodes(2).Text = sFileLen + DFLS
				Case 2
					'Nothing here...
			End Select
		Else
			Select Case NodeIndex
				Case 0
					GetDefaultFileParameters(FileNode, A)
				Case 1
					GetDefaultFileParameters(FileNode, A, O)
				Case 2
					GetDefaultFileParameters(FileNode, A, O, L)
			End Select
		End If

		ValidateFileParameters(FileNode)    'This will make sure parameters are within limits
		'It will also update file size and bundle size(s)

		CheckFileParameterColors(FileNode)  'Check if parameters are default or not

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:
		FinishUpdate()

	End Sub


	Private Sub GetFile(FileName As String)
		On Error GoTo Err

		If IO.File.Exists(Replace(FileName, "*", "")) Then
			P = IO.File.ReadAllBytes(Replace(FileName, "*", ""))
			Ext = LCase(Strings.Right(Replace(FileName, "*", ""), 4))
			PLen = P.Length
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub GetDefaultFileParameters(FileNode As TreeNode, Optional FA As String = "", Optional FO As String = "", Optional FL As String = "")
		On Error GoTo Err

		StartUpdate()

		'Determine default numeric parameters
		Select Case Ext
			Case ".sid"
				DFAN = P(P(7)) + (P(P(7) + 1) * 256)
				DFON = P(7) + 2
			Case Else
				'Default load address = first 2 bytes of file as in .prg files
				If PLen > 2 Then
					DFAN = P(0) + (P(1) * 256)
				Else
					DFAN = 2064 'If file is less than3 bytes long, load address is arbitrary 2064
				End If
				'Default offset depends on load address
				If FA = "" Then
					DFON = 2
				ElseIf DFAN = Convert.ToInt32(FA, 16) Then
					DFON = 2
				Else
					DFON = 0
				End If
		End Select

		'Default length depends on offset
		If FO = "" Then
			DFLN = PLen - DFON
		Else
			DFLN = PLen - Convert.ToInt32(FO, 16)
		End If

		'File Address+ File Length cannot be > $10000
		If FA <> "" Then    'We have a file address
			'Calculate DFLN using current file address
			If Convert.ToInt32(FA, 16) + DFLN > &H10000 Then
				DFLN = &H10000 - Convert.ToInt32(FA, 16)
			End If
		Else
			'Calculate DFLN using default file address
			If DFAN + DFLN > &H10000 Then
				DFLN = &H10000 - DFAN
			End If
		End If

		'If Load Address=$0000 then, limit DFLN to $ffff to make sure it fits in word format
		If DFLN = &H10000 Then DFLN -= 1

		'Calculate default parameter strings
		DFAS = ConvertIntToHex(DFAN, 4)
		DFOS = ConvertIntToHex(DFON, 8)
		DFLS = ConvertIntToHex(DFLN, 4)
		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		FinishUpdate()

	End Sub

	Private Sub ResetFileParameters(FileNode As TreeNode, NodeIndex As Integer)
		On Error GoTo Err

		StartUpdate()

		Select Case NodeIndex
			Case 0
				txtEdit.Text = DFAS
				With FileNode.Nodes(0)
					.Text = sFileAddr + DFAS
					.ForeColor = colFileParamDefault
				End With
				With FileNode.Nodes(1)
					.Text = sFileOffs + DFOS
					.ForeColor = colFileParamDefault
				End With
				With FileNode.Nodes(2)
					.Text = sFileLen + DFLS
					.ForeColor = colFileParamDefault
				End With

			Case 1
				txtEdit.Text = DFOS
				With FileNode.Nodes(1)
					.Text = sFileOffs + DFOS
					.ForeColor = colFileParamDefault
				End With
				With FileNode.Nodes(2)
					.Text = sFileLen + DFLS
					.ForeColor = colFileParamDefault
				End With
			Case 2
				txtEdit.Text = DFLS
				With FileNode.Nodes(2)
					.Text = sFileLen + DFLS
					.ForeColor = colFileParamDefault
				End With

			Case Else
				Exit Select
		End Select

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		FinishUpdate()

	End Sub

	Private Sub ValidateFileParameters(FileNode As TreeNode)
		On Error GoTo Err

		NewFile = FileNode.Text

		FAddr = Convert.ToInt32(Strings.Right(FileNode.Nodes(0).Text, 4), 16)
		FOffs = Convert.ToInt32(Strings.Right(FileNode.Nodes(1).Text, 8), 16)
		FLen = Convert.ToInt32(Strings.Right(FileNode.Nodes(2).Text, 4), 16)

		'If Loading = False Then TV.BeginUpdate()
		StartUpdate()

		'Make sure Offset is within program length
		If FOffs > PLen - 1 Then
			FOffs = PLen - 1
			FileNode.Nodes(1).Text = sFileOffs + ConvertIntToHex(FOffs, 8)
		End If

		'Check file length
		If (FLen = 0) Or (FOffs + FLen > PLen) Then
			FLen = PLen - FOffs
			DFLN = FLen
			DFLS = ConvertIntToHex(DFLN, 4)
			FileNode.Nodes(2).Text = sFileLen + ConvertIntToHex(FLen, 4)
		End If

		'Make sure file is within memory
		If FAddr + FLen > &HFFFF Then
			FLen = &H10000 - FAddr
			FileNode.Nodes(2).Text = sFileLen + ConvertIntToHex(FLen, 4)
		End If

		FAS = ConvertIntToHex(FAddr, 4)
		FOS = ConvertIntToHex(FOffs, 8)
		FLS = ConvertIntToHex(FLen, 4)

		'----------------------------------
		FileNode.Text = CalcFileSize(FileNode.Text, FAddr, FLen)
		FileNode.Nodes(4).Text = sFileSize + FileSize.ToString + " block" + If(FileSize = 1, "", "s")
		'----------------------------------

		With FileNode
			If OverlapsIO() = False Then
				.Text = Replace(.Text, "*", "")
				'.Nodes(3).Text = sFileUIO + "n/a"
				.Nodes(3).Text = sFileUIO + "RAM"
				.Nodes(3).ForeColor = colFileIODefault
			ElseIf InStr(.Text, "*") <> 0 Then
				'.Nodes(3).Text = sFileUIO + "yes"
				.Nodes(3).Text = sFileUIO + "RAM"
				.Nodes(3).ForeColor = colFileIOEdited
			Else
				'.Nodes(3).Text = sFileUIO + " no"
				.Nodes(3).Text = sFileUIO + "I/O"
				.Nodes(3).ForeColor = colFileIODefault
			End If
		End With

		'----------------------------------
		If ChkSize.Checked Then CalcDiskSizeWithForm(BaseNode, FileNode.Parent.Index)
		'----------------------------------

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		FinishUpdate()

	End Sub

	Private Sub CheckFileParameterColors(FileNode As TreeNode)
		On Error GoTo Err

		StartUpdate()

		If Strings.Right(FileNode.Nodes(0).Text, 4) = DFAS Then
			DFA = True
		Else
			DFA = False
		End If

		If Strings.Right(FileNode.Nodes(1).Text, 8) = DFOS Then
			DFO = True
			If DFOS = "00000000" Then
				DFA = False
			End If
		Else
			DFO = False
			DFA = False
		End If

		If Strings.Right(FileNode.Nodes(2).Text, 4) = DFLS Then
			DFL = True
		Else
			DFL = False
			DFO = False
			DFA = False
		End If

		'Update File Parameter Node Colors
		FileNode.Nodes(0).ForeColor = If(DFA = True, colFileParamDefault, colFileParamEdited)
		FileNode.Nodes(1).ForeColor = If(DFO = True, colFileParamDefault, colFileParamEdited)
		FileNode.Nodes(2).ForeColor = If(DFL = True, colFileParamDefault, colFileParamEdited)

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		'If Loading = False Then TV.EndUpdate()
		FinishUpdate()

	End Sub

	Private Function OverlapsIO() As Boolean
		On Error GoTo Err

		If (FAddr >= &HD000) And (FAddr < &HE000) Then
			'File beings under IO
			OverlapsIO = True
		ElseIf (FAddr + FLen - 1 >= &HD000) And (FAddr + FLen - 1 < &HE000) Then
			'File ends under IO
			OverlapsIO = True
		ElseIf (FAddr < &HD000) And (FAddr + FLen - 1 >= &HE000) Then
			'File begins before and ends after IO
			OverlapsIO = True
		Else
			'Otherwise
			OverlapsIO = False
		End If

		Exit Function
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Sub UpdateDiskPath()
		On Error GoTo Err

		Dim N As TreeNode = TV.SelectedNode

		FileType = 3  'D64 file

		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		N.Text = sDiskPath + NewFile

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub UpdateDirArtPath()
		On Error GoTo Err

		Dim N As TreeNode = TV.SelectedNode

		FileType = 1  'Text file

		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		N.Text = sDirArt + NewFile

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub UpdateFileParameters(FileNode As TreeNode)
		On Error GoTo Err

		If FileNode Is Nothing Then Exit Sub

		StartUpdate()

		If FileNode.Nodes(FileNode.Name + ":FA") Is Nothing Then
			FileNode.Nodes.Add(FileNode.Name + ":FA", sFileAddr + ConvertIntToHex(FAddr, 4))
			With FileNode.Nodes(FileNode.Name + ":FA")
				.Tag = FileNode.Tag
				.ForeColor = If(DFA = True, colFileParamDefault, colFileParamEdited)
				.NodeFont = New Font("Consolas", 10)
			End With
		Else
			FileNode.Nodes(FileNode.Name + ":FA").Text = sFileAddr + ConvertIntToHex(FAddr, 4)
		End If

		If FileNode.Nodes(FileNode.Name + ":FO") Is Nothing Then
			FileNode.Nodes.Add(FileNode.Name + ":FO", sFileOffs + ConvertIntToHex(FOffs, 8))
			With FileNode.Nodes(FileNode.Name + ":FO")
				.Tag = FileNode.Tag
				.ForeColor = If(DFO = True, colFileParamDefault, colFileParamEdited)
				.NodeFont = New Font("Consolas", 10)
			End With
		Else
			FileNode.Nodes(FileNode.Name + ":FO").Text = sFileOffs + ConvertIntToHex(FOffs, 8)
		End If

		If FileNode.Nodes(FileNode.Name + ":FL") Is Nothing Then
			FileNode.Nodes.Add(FileNode.Name + ":FL", sFileLen + ConvertIntToHex(FLen, 4))
			With FileNode.Nodes(FileNode.Name + ":FL")
				.Tag = FileNode.Tag
				.ForeColor = If(DFL = True, colFileParamDefault, colFileParamEdited)
				.NodeFont = New Font("Consolas", 10)
			End With
		Else
			FileNode.Nodes(FileNode.Name + ":FL").Text = sFileLen + ConvertIntToHex(FLen, 4)
		End If

		If FileNode.Nodes(FileNode.Name + ":FUIO") Is Nothing Then
			'                                                           Overlaps I/O      AND DOES NOT GO UNDER I/O
			FileNode.Nodes.Add(FileNode.Name + ":FUIO", sFileUIO + If((OverlapsIO() = True) And (FileUnderIO = False), "I/O", "RAM"))
			With FileNode.Nodes(FileNode.Name + ":FUIO")
				.Tag = FileNode.Tag
				.ForeColor = If(FileUnderIO = True, colFileIOEdited, colFileIODefault)
				.NodeFont = New Font("Consolas", 10)
			End With
		Else
			FileNode.Nodes(FileNode.Name + ":FUIO").Text = sFileUIO + If((OverlapsIO() = True) And (FileUnderIO = False), "I/O", "RAM")
		End If

		If FileNode.Nodes(FileNode.Name + ":FS") Is Nothing Then
			FileNode.Nodes.Add(FileNode.Name + ":FS", sFileSize + FileSize.ToString + " block" + If(FileSize <> 1, "s", ""))
			With FileNode.Nodes(FileNode.Name + ":FS")
				.Tag = FileNode.Tag
				.ForeColor = colFileSize
				.NodeFont = New Font("Consolas", 10)
			End With
		Else
			FileNode.Nodes(FileNode.Name + ":FS").Text = sFileSize + FileSize.ToString + " block" + If(FileSize <> 1, "s", "")
		End If

		FinishUpdate()

		FileNode.Expand()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Function CalcFileSize(FN As String, FA As Integer, FL As Integer) As String
		On Error GoTo Err

		CalcFileSize = FN

		Dim DefaultFUIO As Boolean = InStr(FN, "*") <> 0

		FAddr = FA
		FLen = FL

		FileSize = CalcOrigBlockCnt()   'This also opens the prg to Prg() and calculates FAddr and FLen

		FN = Replace(FN, "*", "")

		If OverlapsIO() = True Then
			If DefaultFUIO = False Then
				FileUnderIO = False
			Else
				FN += "*"
				FileUnderIO = True
			End If
		Else
			FileUnderIO = False
		End If

		CalcFileSize = FN

		Exit Function
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Function CalcOrigBlockCnt() As Integer
		On Error GoTo Err

		CalcOrigBlockCnt = 0

		If NewFile = "" Then Exit Function

		ReDim Prg(0)

		If Strings.Right(NewFile, 1) = "*" Then
			'Prg = IO.File.ReadAllBytes(Strings.Left(NewFile, Strings.Len(NewFile) - 1))
			Prg = IO.File.ReadAllBytes(Replace(NewFile, "*", ""))
			Ext = LCase(Strings.Right(Replace(NewFile, "*", ""), 3))
			FileUnderIO = True
		Else
			Prg = IO.File.ReadAllBytes(NewFile)
			Ext = LCase(Strings.Right(NewFile, 3))
			FileUnderIO = False
		End If

		DFA = True
		DFO = True
		DFL = True
		Select Case Ext
			Case "sid"
				If FAddr = -1 Then
					FAddr = Prg(Prg(7)) + (Prg(Prg(7) + 1) * 256)
				Else
					If FAddr <> Prg(Prg(7)) + (Prg(Prg(7) + 1) * 256) Then
						DFA = False
					End If
				End If
				If FOffs = -1 Then
					FOffs = Prg(7) + 2
				Else
					If FOffs <> Prg(7) + 2 Then
						DFO = False
						DFA = False
					End If
				End If
				If FLen = 0 Then
					FLen = Prg.Length - FOffs
				Else
					If FLen <> Prg.Length - FOffs Then
						DFL = False
						DFO = False
						DFA = False
					End If
				End If
			Case Else   'including prg: load address derived from first 2 bytes, offset=2, length=prg length-2
				If FAddr = -1 Then
					FAddr = Prg(0) + (Prg(1) * 256)
				Else
					If FAddr <> Prg(0) + (Prg(1) * 256) Then
						DFA = False
					End If
				End If
				If FOffs = -1 Then
					FOffs = 2
				Else
					If FOffs <> 2 Then
						DFO = False
						DFA = False
					End If
				End If
				If FLen = 0 Then
					FLen = Prg.Length - FOffs
				Else
					If (FLen <> Prg.Length - FOffs) And (FLen <> &H10000 - FAddr) Then
						DFL = False
						DFO = False
						DFA = False
					End If
				End If
		End Select

		CalcOrigBlockCnt = Int(FLen / 254)
		If FLen Mod 254 <> 0 Then CalcOrigBlockCnt += 1

		Exit Function
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Sub BtnLoad_Click(sender As Object, e As EventArgs) Handles BtnLoad.Click
		On Error GoTo Err

		OpenFile("Sparkle Loader Script files", "Sparkle Loader Script files (*.sls)|*.sls")
		If NewFile <> "" Then

			SetScriptPath(NewFile)

			OpenScript()

			NodeSelect()

			SelNode.EnsureVisible()
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
		On Error GoTo Err

		ConvertNodesToScript()

		If Script = "" Then
			Exit Sub
		End If

		SaveFile("Save Sparkle Loader Script files")

		If NewFile <> "" Then
			BaseNode.Text = "This Script: " + NewFile

			SetScriptPath(NewFile)

			If OptRelativePaths.Checked = True Then
				Script = Replace(Script, ScriptPath, "")
			End If

			tssLabel.Text = "Script: " + ScriptName
			IO.File.WriteAllText(ScriptName, Script)
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub UpdateFileNode()
		On Error GoTo Err

		Dim N As TreeNode = TV.SelectedNode

		FileType = 2  'Prg file

		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		StartUpdate()

		FAddr = -1
		FOffs = -1
		FLen = 0

		If N.Index = 0 Then
			BlockCnt = 0
		Else
			BlockCnt = 1    'Fake block for compression
		End If

		N.Text = NewFile
		N.NodeFont = New Font(TV.Font.Name, TV.Font.Size, FontStyle.Regular)

		GetFile(N.Text)
		GetDefaultFileParameters(N)

		FAddr = DFAN
		FOffs = DFON
		FLen = DFLN
		DFA = True
		DFO = True
		DFL = True

		N.Text = CalcFileSize(N.Text, DFAN, DFLN)

		UpdateFileParameters(N)
		CheckFileParameterColors(N)

		CalcDiskSizeWithForm(BaseNode, N.Parent.Index)

		FinishUpdate()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles BtnOK.Click
		On Error GoTo Err

		bBuildDisk = True

		Me.Close()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub SwapNewBlockStatus()
		On Error GoTo Err

		With SelNode
			If Strings.Right(SelNode.Text, 2) = "NO" Then
				.Text = sNewBlock + "YES"
				.ForeColor = colNewBlockYes
			Else
				.Text = sNewBlock + "NO"
				.ForeColor = colNewBlockNo
			End If
		End With

		'Dim Frm As New FrmDisk
		'Frm.Show(Me)

		'StartUpdate()
		If SelNode.Parent.Nodes.Count > 2 Then
			CalcDiskSizeWithForm(BaseNode, SelNode.Parent.Index)
		End If
		'FinishUpdate()

		'Frm.Close()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnEntryUp_Click(sender As Object, e As EventArgs) Handles BtnEntryUp.Click
		On Error GoTo Err

		If TV.SelectedNode Is Nothing Then Exit Sub

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		Dim N As TreeNode = TV.SelectedNode
		Dim I As Integer = N.Index

		If I > 0 Then
			StartUpdate()
			BaseNode.Nodes.RemoveAt(I)
			BaseNode.Nodes.Insert(I - 1, N)
			'TV.Refresh()
			CalcDiskSize(BaseNode, I - 1)
			FinishUpdate()
			TV.SelectedNode = N
			TV.Focus()
		End If

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:
		Frm.Close()

	End Sub

	Private Sub BtnEntryDown_Click(sender As Object, e As EventArgs) Handles BtnEntryDown.Click
		On Error GoTo Err

		If TV.SelectedNode Is Nothing Then Exit Sub

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		Dim N As TreeNode = TV.SelectedNode
		Dim I As Integer = N.Index

		If I < BaseNode.Nodes.Count - 2 Then
			StartUpdate()
			BaseNode.Nodes.RemoveAt(I)
			BaseNode.Nodes.Insert(I + 1, N)
			'TV.Refresh()
			CalcDiskSize(BaseNode, I)
			FinishUpdate()
			TV.SelectedNode = N
			TV.Focus()
		End If

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		Frm.Close()

	End Sub

	Private Sub SwapIOStatus()
		On Error GoTo Err

		With SelNode
			If Strings.Right(.Text, 3) = "I/O" Then
				.Text = sFileUIO + "RAM"
				.ForeColor = colFileIOEdited
				.Parent.Text = Replace(.Parent.Text, "*", "") + "*"
				'----------------------------
				CalcDiskSizeWithForm(BaseNode, .Parent.Parent.Index)
				'----------------------------
			Else    'Strings.Right(.Text, 3)="RAM"
				FAddr = Convert.ToInt32(Strings.Right(.Parent.Nodes(0).Text, 4), 16)
				FLen = Convert.ToInt32(Strings.Right(.Parent.Nodes(2).Text, 4), 16)
				If OverlapsIO() Then    'Only change status if file overlaps I/O
					.Text = sFileUIO + "I/O"
					.ForeColor = colFileIODefault
					.Parent.Text = Replace(.Parent.Text, "*", "")
					'----------------------------
					CalcDiskSizeWithForm(BaseNode, .Parent.Parent.Index)
					'----------------------------
				End If
			End If
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ChkSize_CheckedChanged(sender As Object, e As EventArgs) Handles ChkSize.CheckedChanged
		On Error GoTo Err

		CalcDiskSizeWithForm(BaseNode)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub OpenScript()
		On Error GoTo Err

		tssLabel.Text = "Script: " + ScriptName

		Script = IO.File.ReadAllText(ScriptName)

		'----------------------
		ConvertScriptToNodes()
		'----------------------

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnNew_Click(sender As Object, e As EventArgs) Handles BtnNew.Click
		On Error GoTo Err

		If MsgBox("Do you really want to start a new script?", vbQuestion + vbYesNo + vbDefaultButton2, "New script?") = vbNo Then Exit Sub

		SetScriptPath("")

		tssLabel.Text = "Script: (New Script)"

		ResetArrays()

		SC = 0
		DC = 0
		PC = 0
		FC = 0

		CurrentScript = SC
		CurrentDisk = DC
		CurrentBundle = PC
		CurrentFile = FC

		ReDim BundleSizeA(PC), BundleBytePtrA(PC), BundleBitPtrA(PC), BundleBitPosA(PC), BundleOrigSizeA(PC)

		BundleBytePtrA(PC) = 255
		BundleBitPtrA(PC) = 0
		BundleBitPosA(PC) = 15

		DiskCnt = DC
		ReDim DiskSizeA(DiskCnt)

		StartUpdate()

		TV.Nodes.Clear()

		LoopSet = False
		ZPSet = False

		AddBaseNode()

		NewD = False

		FinishUpdate()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_NodeMouseHover(sender As Object, e As TreeNodeMouseHoverEventArgs) Handles TV.NodeMouseHover
		On Error GoTo Err

		TT.Hide(TV)

		If ChkToolTips.Checked = False Then Exit Sub

		If Loading = True Then Exit Sub

		Dim TTT As String = ""

		With TT
			If e.Node.Name = NewDiskKey Then
				.ToolTipTitle = "Add New Demo Disk"
				TTT = tAddDisk
			ElseIf e.Node.Name = BaseScriptKey Then
				.ToolTipTitle = "This script"
				TTT = tBaseScript
			ElseIf e.Node.Name = NewBundleKey Then
				.ToolTipTitle = "Add New File Bunlde "
				TTT = tAddBundle
			ElseIf e.Node.Name = NewScriptKey Then
				.ToolTipTitle = "Add New Script"
				TTT = tAddScript
			ElseIf Strings.Right(e.Node.Name, 7) = NewFileKey Then
				.ToolTipTitle = "Add New Demo File"
				TTT = tAddFile
			ElseIf Strings.Right(e.Node.Name, 3) = ":NB" Then
				.ToolTipTitle = "Align Bundle with Disk Sector"
				TTT = tNB
			ElseIf Strings.Right(e.Node.Name, 3) = ":FS" Then
				.ToolTipTitle = "File Size"
				TTT = tFileSize
			ElseIf Strings.Right(e.Node.Name, 3) = ":FA" Then
				.ToolTipTitle = "File Address"
				TTT = tFileAddr
			ElseIf Strings.Right(e.Node.Name, 3) = ":FO" Then
				.ToolTipTitle = "File Offset"
				TTT = tFileOffs
			ElseIf Strings.Right(e.Node.Name, 3) = ":FL" Then
				.ToolTipTitle = "File Length"
				TTT = tFileLen
			ElseIf Strings.Right(e.Node.Name, 5) = ":FUIO" Then
				.ToolTipTitle = "I/O Status"
				TTT = tLoadUIO
			ElseIf Strings.InStr(e.Node.Name, ":F") > 0 Then
				.ToolTipTitle = "Demo File"
				TTT = tFile
			ElseIf Strings.Left(e.Node.Text, 5) = "[Disk" Then
				.ToolTipTitle = "Demo Disk"
				TTT = tDisk
			ElseIf Strings.Left(e.Node.Text, 5) = "[Bund" Then
				.ToolTipTitle = "File Bundle"
				TTT = tBundle
			ElseIf Strings.Left(e.Node.Text, 7) = "[Script" Then
				.ToolTipTitle = "Embedded script"
				TTT = tScript
			Else
				Dim S As String = Strings.Left(e.Node.Name, InStr(e.Node.Name, ":") + 1)
				Select Case S
					Case sDiskPath
						.ToolTipTitle = "Disk Path"
						TTT = tDiskPath
					Case sDiskHeader
						.ToolTipTitle = "Disk Header"
						TTT = tDiskHeader
					Case sDiskID
						.ToolTipTitle = "Disk ID"
						TTT = tDiskID
					Case sDemoName
						.ToolTipTitle = "Demo Name"
						TTT = tDemoName
					Case sDemoStart
						.ToolTipTitle = "Demo Start"
						TTT = tDemoStart
					Case sDirArt
						.ToolTipTitle = "DirtArt"
						TTT = tDirArt
					Case sZP
						.ToolTipTitle = "Zeropage Usage"
						TTT = tZP
						'Case sPacker
						'.ToolTipTitle = "Packer to be used"
						'TTT = tPacker
					Case sLoop
						.ToolTipTitle = "Looping after the last disk"
						TTT = tLoop
					Case Else
				End Select
			End If

			If TTT <> "" Then
				.Show(TTT, TV)
			End If
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
		On Error GoTo Err

		bBuildDisk = False

		Me.Close()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub UpdateScriptPath()
		On Error GoTo Err

		Dim N As TreeNode = TV.SelectedNode

		FileType = 5  'SLS file
		OpenDemoFile()

		If NewFile = "" Then Exit Sub

		N.Text = BaseScriptText + NewFile

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ChkExpand_CheckedChanged(sender As Object, e As EventArgs) Handles ChkExpand.CheckedChanged
		On Error GoTo Err

		If TV.Nodes.Count > 0 Then
			ToggleFileNodes(BaseNode)
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub
	Private Sub ResetBundleArrays()

		ReDim BundleSizeA(PC), BundleBytePtrA(PC), BundleBitPtrA(PC), BundleBitPosA(PC), BundleOrigSizeA(PC)

		BundleBytePtrA(PC) = 255
		BundleBitPtrA(PC) = 0
		BundleBitPosA(PC) = 15

	End Sub

	Private Sub ConvertScriptToNodes()
		On Error GoTo Err

		If Script = "" Then Exit Sub

		Dim ScriptPath As String = ""

		For I As Integer = Len(ScriptName) To 1 Step -1
			If Mid(ScriptName, I, 1) = "\" Then
				ScriptPath = Strings.Left(ScriptName, I)            'Path
				Exit For
			End If
		Next

		Dim Lines() As String = Script.Split(vbLf)           'Split script to lines at VbLf (this may leave VbCr at line endings)

		If Lines(0).TrimEnd(Chr(13)) <> ScriptHeader Then
			MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
			Exit Sub
		End If

		StartUpdate()
		Loading = True

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		'RESET ALL COUNTERS
		DC = 0
		PC = 0
		FC = 0
		SC = 0

		CurrentDisk = DC
		CurrentBundle = PC
		CurrentFile = FC

		ResetBundleArrays()

		BaseNode.Text = "This Script: " + ScriptName

		BaseNode.Nodes.Clear()

		LoopSet = False
		ZPSet = False

		Dim Fnt As New Font("Consolas", 10)
		NewBundle = True
		NewD = False
		For I As Integer = 1 To Lines.Count - 1

			Lines(I) = Lines(I).TrimEnd(Chr(13))        'If VbCrLF was used then remove VbCr from line endings

			If InStr(Lines(I), vbTab) = 0 Then
				ScriptEntryType = Lines(I)
			Else
				ScriptEntryType = Strings.Left(Lines(I), InStr(Lines(I), vbTab) - 1)
				ScriptEntry = Strings.Right(Lines(I), Len(Lines(I)) - InStr(Lines(I), vbTab)).TrimStart(vbTab)
			End If

			SplitEntry()

			Select Case LCase(ScriptEntryType)
				Case ""     'New Bundle
					NewBundle = True
				Case "path:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
					If DiskNode.Nodes(sDiskPath + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDiskPath + DC.ToString, sDiskPath + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDiskPath + DC.ToString), sDiskPath + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDiskPath)
					End If
					NewBundle = True
				Case "header:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If DiskNode.Nodes(sDiskHeader + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDiskHeader + DC.ToString, sDiskHeader + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDiskHeader + DC.ToString), sDiskHeader + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					End If
					NewBundle = True
				Case "id:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If DiskNode.Nodes(sDiskID + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDiskID + DC.ToString, sDiskID + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDiskID + DC.ToString), sDiskID + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDiskID)
					End If
					NewBundle = True
				Case "name:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If DiskNode.Nodes(sDemoName + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDemoName + DC.ToString, sDemoName + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDemoName + DC.ToString), sDemoName + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDemoName)
					End If
					NewBundle = True
				Case "start:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If DiskNode.Nodes(sDemoStart + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDemoStart + DC.ToString, sDemoStart + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDemoStart + DC.ToString), sDemoStart + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfo, Fnt) ', tDemoStart)
					End If
					NewBundle = True
				Case "dirart:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If ScriptEntryArray(0) <> "" Then
						If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
					End If
					If DiskNode.Nodes(sDirArt + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDirArt + DC.ToString, sDirArt + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDirArt + DC.ToString), sDirArt + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					End If
					NewBundle = True
					'Case "packer:"
					'If NewD = False Then
					'NewD = True
					'AddDiskToScriptNode(BaseNode)
					'End If
					'If DiskNode.Nodes(sPacker + DC.ToString) Is Nothing Then
					'AddNode(DiskNode, sPacker + DC.ToString, sPacker + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfo, Fnt)
					'Else
					'UpdateNode(DiskNode.Nodes(sPacker + DC.ToString), sPacker + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfo, Fnt)
					'End If
					'Select Case LCase(ScriptEntryArray(0))
					'Case "faster"
					'Packer = 1
					'Case Else   '"better"
					'Packer = 2
					'End Select
					'NewPart = True
				Case "zp:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					'Correct length
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("02", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					UpdateNode(ZPNode, sZP + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfo, Fnt)
					If ZPSet = False Then 'ZP can only be set from the first disk
						DiskNode.Nodes.Add(ZPNode)
						ZPSet = True
					End If
					NewBundle = True
				Case "loop:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					UpdateNode(DiskNode.Nodes(sLoop), sLoop + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					If LoopSet = False Then 'Loop can only be set from the first disk
						DiskNode.Nodes.Add(LoopNode)
						LoopSet = True
					End If
					NewBundle = True
				Case "il0:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("04", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL0 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL0 + DC.ToString, sIL0 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL0 + DC.ToString), sIL0 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "il1:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("03", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL1 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL1 + DC.ToString, sIL1 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL1 + DC.ToString), sIL1 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "il2:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("03", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL2 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL2 + DC.ToString, sIL2 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL2 + DC.ToString), sIL2 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "il3:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("03", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL3 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL3 + DC.ToString, sIL3 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL3 + DC.ToString), sIL3 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfo, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "list:", "script:"
					NewD = False
					If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)

					AddScriptToScriptNode(BaseNode, ScriptEntryArray(0))
					ConvertScriptToScriptNodes(ScriptNode, ScriptEntryArray(0))

				Case "file:"
					NewD = False
					If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)

					CorrectFileParameterFormat()
					AddFileToScriptNode(BaseNode, ScriptEntryArray(0), If(ScriptEntryArray.Count > 1, ScriptEntryArray(1), ""), If(ScriptEntryArray.Count > 2, ScriptEntryArray(2), ""), If(ScriptEntryArray.Count > 3, ScriptEntryArray(3), ""))
					NewBundle = False
				Case "new block", "next block", "new sector", "align bundle", "align"
					BundleInNewBlock = True
				Case Else
					'Figure out what to do with comments here...
			End Select
		Next

		AddNewEntryNode()

		'MsgBox(DC.ToString + vbNewLine + PC.ToString + vbNewLine + FC.ToString + vbNewLine + SC.ToString)

		CalcDiskSize(BaseNode)

		ToggleFileNodes(BaseNode)

		GoTo Done

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		Frm.Close()

		Loading = False
		FinishUpdate()
		BaseNode.Expand()

	End Sub

	Private Sub ConvertScriptToScriptNodes(SN As TreeNode, SPath As String)
		On Error GoTo Err

		If IO.File.Exists(SPath) = False Then Exit Sub

		Dim S As String = IO.File.ReadAllText(SPath)

		If S = "" Then Exit Sub

		Dim Path As String = ""
		For I As Integer = Len(SPath) To 1 Step -1
			If Mid(SPath, I, 1) = "\" Then
				Path = Strings.Left(SPath, I)            'Path
				Exit For
			End If
		Next

		'If InStr(SPath, "\") = 0 Then Exit Sub

		SN.Nodes.Clear()

		Dim Lines() As String = S.Split(vbLf)           'Split script to lines at VbLf (this may leave VbCr at line endings)

		'If SN.Name = BaseNode.Name  Then                       'Check if this is a valid script file only if this is the master script
		'If Lines(0).TrimEnd(Chr(13)) <> ScriptHeader Then
		'MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
		'Exit Sub
		'End If
		'End If

		'Empty line (new bundle) will be a single $0A charachter, otherwise, line string will not contain vbNewLine

		Dim Fnt As New Font("Consolas", 10)
		NewBundle = True
		NewD = False
		For I As Integer = 0 To Lines.Count - 1

			Lines(I) = Lines(I).TrimEnd(Chr(13))        'If VbCrLF was used then remove VbCr from line endings

			If InStr(Lines(I), vbTab) = 0 Then
				ScriptEntryType = Lines(I)
			Else
				ScriptEntryType = Strings.Left(Lines(I), InStr(Lines(I), vbTab) - 1)
				ScriptEntry = Strings.Right(Lines(I), Len(Lines(I)) - InStr(Lines(I), vbTab)).TrimStart(vbTab)
			End If

			SplitEntry()

			Select Case LCase(ScriptEntryType)
				Case ""     'New Bundle
					NewBundle = True
				Case "path:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = Path + ScriptEntryArray(0)
					If DiskNode.Nodes(sDiskPath + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDiskPath + DC.ToString, sDiskPath + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDiskPath + DC.ToString), sDiskPath + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDiskPath)
					End If
					NewBundle = True
				Case "header:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					If DiskNode.Nodes(sDiskHeader + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDiskHeader + DC.ToString, sDiskHeader + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDiskHeader + DC.ToString), sDiskHeader + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					End If
					NewBundle = True
				Case "id:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					If DiskNode.Nodes(sDiskID + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDiskID + DC.ToString, sDiskID + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDiskID + DC.ToString), sDiskID + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDiskID)
					End If
					NewBundle = True
				Case "name:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					If DiskNode.Nodes(sDemoName + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDemoName + DC.ToString, sDemoName + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDemoName + DC.ToString), sDemoName + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDemoName)
					End If
					NewBundle = True
				Case "start:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					If DiskNode.Nodes(sDemoStart + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDemoStart + DC.ToString, sDemoStart + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDemoStart + DC.ToString), sDemoStart + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDemoStart)
					End If
					NewBundle = True
				Case "dirart:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					If ScriptEntryArray(0) <> "" Then
						If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = Path + ScriptEntryArray(0)
					End If
					If DiskNode.Nodes(sDirArt + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sDirArt + DC.ToString, sDirArt + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sDirArt + DC.ToString), sDirArt + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					End If
					NewBundle = True
					'Case "packer:"
					'If NewD = False Then
					'NewD = True
					'AddDiskToScriptNode(SN)
					'End If
					'If DiskNode.Nodes(sPacker + DC.ToString) Is Nothing Then
					'AddNode(DiskNode, sPacker + DC.ToString, sPacker + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfoGray, Fnt)
					'Else
					'UpdateNode(DiskNode.Nodes(sPacker + DC.ToString), sPacker + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfoGray, Fnt)
					'End If
					'Select Case LCase(ScriptEntryArray(0))
					'Case "faster"
					'Packer = 1
					'Case Else   '"better"
					'Packer = 2
					'End Select
					'NewPart = True
				Case "zp:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					'Correct length
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("02", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					UpdateNode(ZPNode, sZP + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, colDiskInfoGray, Fnt)
					If ZPSet = False Then 'ZP can only be set from the first disk
						DiskNode.Nodes.Add(ZPNode)
						ZPSet = True
					End If
					NewBundle = True
				Case "loop:"
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(SN)
					End If
					UpdateNode(LoopNode, sLoop + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					If LoopSet = False Then 'Loop can only be set from the first disk
						DiskNode.Nodes.Add(LoopNode)
						LoopSet = True
					End If
					NewBundle = True
				Case "il0:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("04", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL0 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL0 + DC.ToString, sIL0 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL0 + DC.ToString), sIL0 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "il1:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("03", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If

					If DiskNode.Nodes(sIL1 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL1 + DC.ToString, sIL1 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL1 + DC.ToString), sIL1 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "il2:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("03", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL2 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL2 + DC.ToString, sIL2 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL2 + DC.ToString), sIL2 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "il3:"
					'If CustomIL Then
					If NewD = False Then
						NewD = True
						AddDiskToScriptNode(BaseNode)
					End If
					If Len(ScriptEntryArray(0)) < 2 Then
						ScriptEntryArray(0) = Strings.Left("03", 2 - Len(ScriptEntryArray(0))) + ScriptEntryArray(0)
					ElseIf Len(ScriptEntryArray(0)) > 2 Then
						ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
					End If
					If DiskNode.Nodes(sIL3 + DC.ToString) Is Nothing Then
						AddNode(DiskNode, sIL3 + DC.ToString, sIL3 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt)
					Else
						UpdateNode(DiskNode.Nodes(sIL3 + DC.ToString), sIL3 + ScriptEntryArray(0), DiskNode.Tag, colDiskInfoGray, Fnt) ', tDemoName)
					End If
					'End If
					NewBundle = True
				Case "list:", "script:"
					NewD = False
					If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = Path + ScriptEntryArray(0)

					AddScriptToScriptNode(SN, ScriptEntryArray(0))
					ConvertScriptToScriptNodes(ScriptNode, ScriptEntryArray(0))
				Case "file:"
					NewD = False
					If InStr(ScriptEntryArray(0), ":") = 0 Then ScriptEntryArray(0) = Path + ScriptEntryArray(0)

					CorrectFileParameterFormat()
					AddFileToScriptNode(SN, ScriptEntryArray(0), If(ScriptEntryArray.Count > 1, ScriptEntryArray(1), ""), If(ScriptEntryArray.Count > 2, ScriptEntryArray(2), ""), If(ScriptEntryArray.Count > 3, ScriptEntryArray(3), ""))
					NewBundle = False
				Case "new block", "next block", "new sector", "align bundle", "align"
					BundleInNewBlock = True
				Case Else
					'Figure out what to do with comments here...
			End Select
		Next

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub
	Private Sub AddDiskToScriptNode(SN As TreeNode)
		On Error GoTo Err

		Dim SNisBaseNode As Boolean = SN.Name = BaseNode.Name

		DC += 1

		CurrentDisk = DC
		AddNode(SN, "D" + DC.ToString, "[Disk" + DC.ToString + "]", DC + &H10000, If(SNisBaseNode, colDisk, colDiskGray))
		DiskNode = SN.Nodes("D" + DC.ToString)
		DiskNode.NodeFont = New Font(TV.Font, FontStyle.Bold)
		If SNisBaseNode Then
			'Dim N As TreeNode = SN.Nodes("D" + DC.ToString)

			Dim Fnt As New Font("Consolas", 10)

			AddNode(DiskNode, sDiskPath + DC.ToString, sDiskPath + "demo.d64", DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sDiskHeader + DC.ToString, sDiskHeader + "demo disk " + Year(Now).ToString, DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sDiskID + DC.ToString, sDiskID + "sprkl", DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sDemoName + DC.ToString, sDemoName + "demo", DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sDemoStart + DC.ToString, sDemoStart + "$", DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sDirArt + DC.ToString, sDirArt, DiskNode.Tag, colDiskInfo, Fnt)
			'AddNode(DiskNode, sPacker + DC.ToString, sPacker + If(My.Settings.DefaultPacker = 1, "faster", "better"), DiskNode.Tag, colDiskInfo, Fnt)
			'If CustomIL Then
			AddNode(DiskNode, sIL0 + DC.ToString, sIL0 + Strings.Left("04", 2 - Len(IL0.ToString)) + IL0.ToString, DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sIL1 + DC.ToString, sIL1 + Strings.Left("03", 2 - Len(IL1.ToString)) + IL1.ToString, DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sIL2 + DC.ToString, sIL2 + Strings.Left("03", 2 - Len(IL2.ToString)) + IL2.ToString, DiskNode.Tag, colDiskInfo, Fnt)
			AddNode(DiskNode, sIL3 + DC.ToString, sIL3 + Strings.Left("03", 2 - Len(IL3.ToString)) + IL3.ToString, DiskNode.Tag, colDiskInfo, Fnt)
			'End If
			If ZPSet = False Then
				ZPNode.Tag = DiskNode.Tag
				DiskNode.Nodes.Add(ZPNode)
				DiskNode.Nodes(sZP).ForeColor = colDiskInfo
				ZPSet = True
			End If
			If LoopSet = False Then
				LoopNode.Tag = DiskNode.Tag
				DiskNode.Nodes.Add(LoopNode)
				DiskNode.Nodes(sLoop).ForeColor = colDiskInfo
				LoopSet = True
			End If
		End If


		'CalcDiskNodeSize(N)

		'AddDiskToScriptNode = DiskNode

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub UpdateNode(Node As TreeNode, Text As String, Optional Tag As Integer = 0, Optional NodeColor As Color = Nothing, Optional NodeFnt As Font = Nothing)
		On Error GoTo Err

		If Node Is Nothing Then Exit Sub

		With Node
			.Text = Text
			.Tag = Tag
			.ForeColor = NodeColor
			.NodeFont = NodeFnt
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddScriptToScriptNode(SN As TreeNode, S As String)
		On Error GoTo Err

		SC += 1

		AddNode(SN, "S" + SC.ToString, "[Script: " + S + "]", &H30000 + SC, If(SN.Name = BaseNode.Name, colScript, colScriptGray))
		ScriptNode = SN.Nodes("S" + SC.ToString)

		If IO.File.Exists(S) = False Then
			ScriptNode.NodeFont = New Font(TV.Font.Name, TV.Font.Size, FontStyle.Italic)
			MsgBox("The following script does not exist:" + vbNewLine + vbNewLine + S, vbOKOnly + vbCritical, "Script cannot be found")
			Exit Sub
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub CorrectFileParameterFormat()
		On Error GoTo Err

		'Correct file parameter lengths to 4-8 characters
		For I As Integer = 1 To ScriptEntryArray.Count - 1

			ScriptEntryArray(I) = Strings.LCase(ScriptEntryArray(I))

			'Remove HEX prefix
			'If InStr(ScriptEntryArray(I), "$") <> 0 Then        'C64
			Replace(ScriptEntryArray(I), "$", "")
			'ElseIf InStr(ScriptEntryArray(I), "&h") <> 0 Then   'VB
			Replace(ScriptEntryArray(I), "&h", "")
			'ElseIf InStr(ScriptEntryArray(I), "0x") <> 0 Then   'C, C++, C#, Java, Python, etc.
			Replace(ScriptEntryArray(I), "0x", "")
			'End If

			'Remove unwanted spaces
			Replace(ScriptEntryArray(I), " ", "")

			Select Case I
				Case 2      'File Offset max. $ffff ffff (dword)
					If ScriptEntryArray(I).Length < 8 Then
						ScriptEntryArray(I) = Strings.Left("00000000", 8 - Len(ScriptEntryArray(I))) + ScriptEntryArray(I)
					ElseIf (I = 2) And (ScriptEntryArray(I).Length > 8) Then
						ScriptEntryArray(I) = Strings.Right(ScriptEntryArray(I), 8)
					End If
				Case Else   'File Address, File Length max. $ffff
					If ScriptEntryArray(I).Length < 4 Then
						ScriptEntryArray(I) = Strings.Left("0000", 4 - Len(ScriptEntryArray(I))) + ScriptEntryArray(I)
					ElseIf ScriptEntryArray(I).Length > 4 Then
						ScriptEntryArray(I) = Strings.Right(ScriptEntryArray(I), 4)
					End If
			End Select
		Next

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddFileToScriptNode(SN As TreeNode, FileName As String, Optional FA As String = "", Optional FO As String = "", Optional FL As String = "")
		On Error GoTo Err

		If NewBundle = True Then AddBundleToScriptNode(SN)

		Dim SNisBaseNode As Boolean = SN.Name = BaseNode.Name

		NewFile = FileName
		GetFile(FileName)

		If FA <> "" Then
			If Len(FA) < 4 Then
				FA = Strings.Left("0000", 4 - Len(FA)) + FA
			End If
		End If

		If FO <> "" Then
			If Len(FO) < 8 Then
				FO = Strings.Left("00000000", 8 - Len(FO)) + FO
			End If
		End If

		If FL <> "" Then
			If Len(FL) < 4 Then
				FL = Strings.Left("0000", 4 - Len(FL)) + FL
			End If
		End If

		FC += 1
		AddNode(BundleNode, BundleNode.Name + ":F" + FC.ToString, FileName, FC, If(SNisBaseNode, colFile, colFileGray))
		FileNode = BundleNode.Nodes(BundleNode.Name + ":F" + FC.ToString)

		If IO.File.Exists(Replace(FileName, "*", "")) = False Then
			FileNode.NodeFont = New Font(TV.Font, FontStyle.Italic)
			MsgBox("The following file does not exist:" + vbNewLine + vbNewLine + FileName, vbOKOnly + vbCritical, "File cannot be found")
			GoTo Done
			'If SNisBaseNode Then AddNewFileEntryNode(PartNode)
			'Exit Sub
		End If

		GetDefaultFileParameters(FileNode, FA, FO, FL)

		Dim FilePath As String = ScriptEntryArray(0)

		If InStr(FilePath, ":") = 0 Then
			FilePath = ScriptPath + FilePath
		End If

		ReDim Prg(0)

		If Strings.Right(FilePath, 1) = "*" Then
			Prg = IO.File.ReadAllBytes(Replace(FilePath, "*", ""))
			Ext = LCase(Strings.Right(Replace(FilePath, "*", ""), 3))
			FileUnderIO = True
		Else
			Prg = IO.File.ReadAllBytes(FilePath)
			Ext = LCase(Strings.Right(FilePath, 3))
			FileUnderIO = False
		End If

		Select Case ScriptEntryArray.Count
			Case 1      'No file parameters in script, use default parameters
				DFA = True
				DFO = True
				DFL = True

				If Ext = "sid" Then                                 'SID file
					FAddr = Prg(Prg(7)) + (Prg(Prg(7) + 1) * 256)
					FOffs = Prg(7) + 2
					FLen = Prg.Length - FOffs
				Else                                                'Any other files
					If Prg.Length > 2 Then                          'We have at least 3 bytes in the file
						FAddr = Prg(0) + (Prg(1) * 256)
						FOffs = 2
					Else                                            'Short file, use arbitrary address
						FAddr = 2064
						FOffs = 0
					End If
					FLen = Prg.Length - FOffs
				End If
			Case 2  'One file parameter = load address
				DFA = False
				DFO = True
				DFL = True
				FAddr = Convert.ToUInt32(ScriptEntryArray(1), 16)   'New File's load address from script
				If Ext = "sid" Then
					If FAddr = Prg(Prg(7)) + (Prg(Prg(7) + 1) * 256) Then DFA = True
					FOffs = Prg(7) + 2
				Else
					If FAddr = Prg(0) + (Prg(1) * 256) Then
						DFA = True
						FOffs = 2
					Else
						FOffs = 0
					End If
				End If
				FLen = Prg.Length - FOffs
			Case 3  'Two file parameters = load address + offset
				DFA = False
				DFO = False
				DFL = True
				FAddr = Convert.ToUInt32(ScriptEntryArray(1), 16)   'New File's load address from script
				FOffs = Convert.ToUInt32(ScriptEntryArray(2), 16)   'New File's offset from script

				If FOffs > Prg.Length - 1 Then                      'Make sure offset is valid
					FOffs = Prg.Length - 1
				End If

				If Ext = "sid" Then
					If FOffs = Prg(7) + 2 Then
						DFO = True
						If FAddr = Prg(Prg(7)) + (Prg(Prg(7) + 1) * 256) Then
							DFA = True
						End If
					End If
				Else
					If FAddr = Prg(0) + (Prg(1) * 256) Then
						If FOffs = 2 Then
							DFA = True
							DFO = True
						End If
					Else
						If FOffs = 0 Then
							DFO = True
						End If
					End If
				End If
				FLen = Prg.Length - FOffs
			Case 4  'All three parameters in script
				DFA = False
				DFO = False
				DFL = False
				FAddr = Convert.ToInt32(ScriptEntryArray(1), 16)   'New File's load address from script
				FOffs = Convert.ToInt32(ScriptEntryArray(2), 16)   'New File's offset from script
				FLen = Convert.ToInt32(ScriptEntryArray(3), 16)    'New File's length from script
				If Ext = "sid" Then
					If FAddr = Prg(Prg(7)) + (Prg(Prg(7) + 1) * 256) Then
						If FOffs = Prg(7) + 2 Then
							If FLen = Prg.Length - FOffs Then
								DFA = True
								DFO = True
								DFL = True
							End If
						Else
							If FLen = Prg.Length - FOffs Then
								DFL = True
							End If
						End If
					Else
						If FOffs = Prg(7) + 2 Then
							If FLen = Prg.Length Then
								DFO = True
								DFL = True
							End If
						Else
							If FLen = Prg.Length Then
								DFL = True
							End If
						End If
					End If
				Else
					If FAddr = Prg(0) + (Prg(1) * 256) Then
						If FOffs = 2 Then
							If FLen = Prg.Length - FOffs Then
								DFA = True
								DFO = True
								DFL = True
							End If
						Else
							If FLen = Prg.Length - FOffs Then
								DFL = True
							End If
						End If
					Else
						If FOffs = 0 Then
							If FLen = Prg.Length Then
								DFO = True
								DFL = True
							End If
						Else
							If FLen = Prg.Length Then
								DFL = True
							End If
						End If
					End If
				End If
		End Select

		If (FLen = 0) Or (FOffs + FLen > Prg.Length) Then   'Make sure length is valid
			FLen = Prg.Length - FOffs
		End If

		If FAddr + FLen > &H10000 Then
			FLen = &H10000 - FAddr
		End If

		FileSize = Int(FLen / 254)
		If FLen Mod 254 <> 0 Then FileSize += 1

		'-----------------------------------------------------------------

		Dim Fnt As New Font("Consolas", 10)
		AddNode(FileNode, FileNode.Name + ":FA", sFileAddr + If(FA <> "", FA, DFAS), FileNode.Tag, If(SNisBaseNode, If(DFA, colFileParamDefault, colFileParamEdited), If(DFA, colFileParamDefaultGray, colFileIOEditedGray)), Fnt)
		AddNode(FileNode, FileNode.Name + ":FO", sFileOffs + If(FO <> "", FO, DFOS), FileNode.Tag, If(SNisBaseNode, If(DFA, colFileParamDefault, colFileParamEdited), If(DFA, colFileParamDefaultGray, colFileIOEditedGray)), Fnt)
		AddNode(FileNode, FileNode.Name + ":FL", sFileLen + If(FL <> "", FL, DFLS), FileNode.Tag, If(SNisBaseNode, If(DFA, colFileParamDefault, colFileParamEdited), If(DFA, colFileParamDefaultGray, colFileIOEditedGray)), Fnt)
		AddNode(FileNode, FileNode.Name + ":FUIO", sFileUIO + If((OverlapsIO() = True) And (FileUnderIO = False), "I/O", "RAM"), FileNode.Tag, If(SNisBaseNode, If(FileUnderIO, colFileIOEdited, colFileIODefault), If(FileUnderIO, colFileIOEditedGray, colFileIODefaultGray)), Fnt)
		AddNode(FileNode, FileNode.Name + ":FS", sFileSize + FileSize.ToString + " block" + If(FileSize = 1, "", "s"), FileNode.Tag, If(SNisBaseNode, colFileSize, colFileSizeGray), Fnt)

		If SNisBaseNode Then
			CheckFileParameterColors(FileNode)  'This will make sure colors are OK
		End If

Done:

		If SNisBaseNode Then
			If BundleNode.Nodes(BundleNode.Name + ":" + NewFileKey) IsNot Nothing Then
				BundleNode.Nodes(BundleNode.Name + ":" + NewFileKey).Remove()
			End If
			AddNewFileEntryNode(BundleNode)
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddBundleToScriptNode(SN As TreeNode)
		On Error GoTo Err
		Dim SNisBaseNode As Boolean = SN.Name = BaseNode.Name

		PC += 1
		ReDim Preserve BundleBytePtrA(PC), BundleBitPtrA(PC), BundleBitPosA(PC), BundleSizeA(PC), BundleOrigSizeA(PC)

		AddNode(SN, "P" + PC.ToString, "[Bundle]", &H20000 + PC, If(SNisBaseNode, colBunlde, colBundleGray))
		BundleNode = SN.Nodes("P" + PC.ToString)
		Dim Fnt As New Font("Consolas", 10)
		AddNode(BundleNode, BundleNode.Name + ":NB", sNewBlock + If(BundleInNewBlock = True, "YES", "NO"), BundleNode.Tag, If(SNisBaseNode, If(BundleInNewBlock = True, colNewBlockYes, colNewBlockNo), If(BundleInNewBlock = True, colNewBlockYesGray, colNewBlockNoGray)), Fnt)
		NewBundle = False
		BundleInNewBlock = False

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ConvertNodesToScript()
		On Error GoTo Err

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		Script = ScriptHeader + vbNewLine

		For I As Integer = 0 To BaseNode.Nodes.Count - 1

			Select Case BaseNode.Nodes(I).Tag
				Case 0 To &HFFFF
					'File nodes cannot be child nodes of the BaseNode
				Case &H10000 To &H1FFFF
					'Disk Node
					Script += vbNewLine
					DiskNode = BaseNode.Nodes(I)
					CurrentDisk = DiskNode.Tag And &HFFFF
					For J As Integer = 0 To DiskNode.Nodes.Count - 1
						Select Case DiskNode.Nodes(J).Name
							Case sDiskPath + CurrentDisk.ToString
								Script += "Path:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sDiskPath)) + vbNewLine
							Case sDiskHeader + CurrentDisk.ToString
								Dim DH As String = Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sDiskHeader))
								If DH = "" Then DH = "demo disk " + Year(Now).ToString
								Script += "Header:" + vbTab + DH + vbNewLine
							Case sDiskID + CurrentDisk.ToString
								Dim DI As String = Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sDiskID))
								If DI = "" Then DI = "sprkl"
								Script += "ID:" + vbTab + DI + vbNewLine
							Case sDemoName + CurrentDisk.ToString
								Script += "Name:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sDemoName)) + vbNewLine
							Case sDemoStart + CurrentDisk.ToString
								Dim SA As String = Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sDemoStart) - 1)
								If SA <> "" Then
									Script += "Start:" + vbTab + SA + vbNewLine
								End If
							Case sDirArt + CurrentDisk.ToString
								Script += "DirArt:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sDirArt)) + vbNewLine
								'Case sPacker + CurrentDisk.ToString
								'Script += "Packer:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sPacker)) + vbNewLine
							Case sZP
								Script += "ZP:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sZP) - 1) + vbNewLine
							Case sLoop
								Script += "Loop:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sLoop)) + vbNewLine
							Case sIL0 + CurrentDisk.ToString
								'If CustomIL Then
								Script += "IL0:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sIL0)) + vbNewLine
								'End If
							Case sIL1 + CurrentDisk.ToString
								'If CustomIL Then
								Script += "IL1:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sIL1)) + vbNewLine
								'End If
							Case sIL2 + CurrentDisk.ToString
								'If CustomIL Then
								Script += "IL2:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sIL2)) + vbNewLine
								'End If
							Case sIL3 + CurrentDisk.ToString
								'If CustomIL Then
								Script += "IL3:" + vbTab + Strings.Right(DiskNode.Nodes(J).Text, Len(DiskNode.Nodes(J).Text) - Len(sIL3)) + vbNewLine
								'End If
						End Select
					Next
				Case &H20000 To &H2FFFF
					'Bundle Node
					BundleNode = BaseNode.Nodes(I)
					Dim NewLineAdded As Boolean = False
					For J As Integer = 0 To BundleNode.Nodes.Count - 1
						If BundleNode.Nodes(J).Tag < &H10000 Then
							If NewLineAdded = False Then
								Script += vbNewLine 'Only add a new line to the script if there are files in the bundle
								NewLineAdded = True 'Do it only once per bundle
								If Strings.Right(LCase(BundleNode.Nodes(0).Text), 3) = "yes" Then
									Script += "Align" + vbNewLine   'We do have a file in this bundle, so check if bundle needs to be aligned with a new sector
								End If
							End If
							FileNode = BundleNode.Nodes(J)
							Script += "File:" + vbTab + FileNode.Text
							If FileNode.Nodes.Count > 0 Then
								If FileNode.Nodes(0).ForeColor = colFileParamEdited Then            'Add Load Address
									Script += vbTab + Strings.Right(FileNode.Nodes(0).Text, 4)
									If FileNode.Nodes(1).ForeColor = colFileParamEdited Then        'Add File Offset
										'Trim Offset down to 4 digits if possible
										Dim O As String = Strings.Right(FileNode.Nodes(1).Text, 8).TrimStart("0")
										If Len(O) < 4 Then O = Strings.Left("0000", 4 - Len(O)) + O
										Script += vbTab + O
										If FileNode.Nodes(2).ForeColor = colFileParamEdited Then    'Add File Length
											Script += vbTab + Strings.Right(FileNode.Nodes(2).Text, 4)
										End If
									End If
								End If
							End If
							Script += vbNewLine
						End If
					Next
				Case &H30000 To &H3FFFF
					'Script Node
					ScriptNode = BaseNode.Nodes(I)
					Dim S As String = ScriptNode.Text.TrimStart("[").TrimEnd("]")
					Script += vbNewLine + "Script:" + vbTab + Strings.Right(S, Len(S) - Len(sScript)) + vbNewLine
				Case Else
			End Select
		Next

		If Replace(Script, vbNewLine, "") = ScriptHeader Then
			Script = ""
		End If

		GoTo Done
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:   Frm.Close()

	End Sub

	Private Sub ReNumberEntries(ParentNode As TreeNode, Optional NodeIndex As Integer = -1)
		On Error GoTo Err

		'THIS WILL RENUMBER ALL DISKS AND PARTS AND COUNT DISKS, PARTS, FILES AND SCRIPTS IN THE SCRIPT

		If ParentNode.Name = BaseNode.Name Then
			DC = 0
			PC = 0
			FC = 0
			SC = 0
			FirstBundleOfDisk = True
			ReDim PDiskNoA(PC), PNewBlockA(PC)
		End If

		If ParentNode.Nodes.Count > 0 Then
			For I As Integer = 0 To ParentNode.Nodes.Count - 1
				Select Case Int(ParentNode.Nodes(I).Tag / &H10000)
					Case 1
						'Disk Node
						DC += 1
						DiskNode = ParentNode.Nodes(I)
						If DiskNode.Index >= NodeIndex Then
							DiskNode.Text = "[Disk " + DC.ToString + "]"
							DiskNode.Name = "D" + DC.ToString
							DiskNode.Tag = &H10000 + DC
						End If
						FirstBundleOfDisk = True
					Case 2
						'Bundle Node
						PC += 1
						ReDim Preserve PDiskNoA(PC), PNewBlockA(PC)
						PDiskNoA(PC) = DC
						BundleNode = ParentNode.Nodes(I)
						PNewBlockA(PC) = Strings.Right(BundleNode.Nodes(0).Text, 3) = "YES"

						If BundleNode.Index >= NodeIndex Then
							BundleNode.Text = "[Bundle " + PC.ToString + "]"
							BundleNode.Name = If(FirstBundleOfDisk, "P", "p") + PC.ToString
							BundleNode.Tag = &H20000 + PC
						End If
						For J As Integer = 1 To BundleNode.Nodes.Count - 1
							If BundleNode.Nodes(J).Tag < &H10000 Then
								FC += 1
							End If
						Next
						FirstBundleOfDisk = False
					Case 3
						'Script Node
						SC += 1
						ReNumberEntries(ParentNode.Nodes(I))
				End Select
			Next
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ScanScript(ParentNode As TreeNode)
		On Error GoTo Err

		'THIS WILL COUNT DISKS, PARTS, FILES AND SCRIPTS IN THE SCRIPT
		'AND RENAME THEM ACCORDING TO ORDER

		If ParentNode.Name = BaseNode.Name Then
			DC = 0
			PC = 0
			FC = 0
			SC = 0
			FirstBundleOfDisk = True
		End If

		If ParentNode.Nodes.Count > 0 Then
			For I As Integer = 0 To ParentNode.Nodes.Count - 1
				Select Case Int(ParentNode.Nodes(I).Tag / &H10000)
					Case 1
						'Disk Node
						DC += 1
						DiskNode = ParentNode.Nodes(I)
						DiskNode.Name = "D" + DC.ToString
						FirstBundleOfDisk = True
					Case 2
						'Bundle Node
						PC += 1
						BundleNode = ParentNode.Nodes(I)
						BundleNode.Name = If(FirstBundleOfDisk, "P", "p") + PC.ToString
						For J As Integer = 1 To BundleNode.Nodes.Count - 1
							If BundleNode.Nodes(J).Tag < &H10000 Then
								FC += 1
							End If
						Next
						FirstBundleOfDisk = False
					Case 3
						'Script Node
						SC += 1
						ScanScript(ParentNode.Nodes(I))
				End Select
			Next
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Function CalcPartSize(PartN As TreeNode, Optional prevPartN As TreeNode = Nothing) As Integer
		On Error GoTo Err

		'THIS WILL CALCULATE THE SIZE OF A SINGLE Bundle, UPDATE Bundle NODE AND RETURN Bundle SIZE

		Dim P() As Byte
		Dim FA, FO, FL As String
		Dim FON, FLN As Integer
		FileCnt = -1
		ReDim FileNameA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt), FileIOA(FileCnt)
		Prgs.Clear()
		ReDim ByteSt(-1)
		UncompBundleSize = 0
		BlockCnt = 0

		For I As Integer = 1 To PartN.Nodes.Count - 1
			If PartN.Nodes(I).Tag < &H10000 Then
				Dim IsItalic As Boolean = False
				If PartN.Nodes(I).NodeFont IsNot Nothing Then
					If PartN.Nodes(I).NodeFont.Italic = True Then
						IsItalic = True
					End If
				End If
				If IsItalic = False Then
					Dim FN As String = PartN.Nodes(I).Text
					Dim FUIO As Boolean

					If InStr(FN, "*") <> 0 Then
						FN = Replace(FN, "*", "")
						FUIO = True
					Else
						FUIO = False
					End If

					If IO.File.Exists(FN) Then
						P = IO.File.ReadAllBytes(FN)
						FA = Strings.Right(PartN.Nodes(I).Nodes(0).Text, 4)
						FO = Strings.Right(PartN.Nodes(I).Nodes(1).Text, 8)
						FL = Strings.Right(PartN.Nodes(I).Nodes(2).Text, 4)

						FON = Convert.ToInt32(FO, 16)
						FLN = Convert.ToInt32(FL, 16)

						If FON + FLN > P.Length Then
							FLN = P.Length - FON
							FL = ConvertIntToHex(FLN, 4)
						End If

						UncompBundleSize += Int(FLN / 254)
						If FLN Mod 254 > 0 Then
							UncompBundleSize += 1
						End If

						Dim PL As List(Of Byte) = P.ToList  'Copy array to list
						P = PL.Skip(FON).Take(FLN).ToArray  'Trim file to specified segment and copy back to array

						FileCnt += 1
						ReDim Preserve FileNameA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt), FileIOA(FileCnt)

						FileNameA(FileCnt) = FN
						FileAddrA(FileCnt) = FA
						FileOffsA(FileCnt) = FO
						FileLenA(FileCnt) = FL
						FileIOA(FileCnt) = FUIO

						Prgs.Add(P)

					Else
						'file does not exist, skip it
						MsgBox("The following file does not exist:" + vbNewLine + vbNewLine + FN, vbOKOnly + vbCritical, "File not found")
						PartN.Nodes(I).NodeFont = New Font(TV.Font, FontStyle.Italic)
					End If
				End If
			End If
		Next

		BundleCnt = CurrentBundle

		BufferCnt = 0
		Dim PrevCP As Integer = 1
		Dim PrevBC As Integer = 0
		Dim PrevPartChgd As Boolean = False
		If (Strings.Left(PartN.Name, 1) = "P") Or (Strings.Right(PartN.Nodes(0).Text, 3) = "YES") Then
			'First bundle on disk or aligned bundle
			ResetBuffer()
		Else

			'Sort bundle, but SortPart Sub works with temporary arrays and liasts

			tmpPrgs = Prgs.ToList
			tmpFileNameA = FileNameA
			tmpFileAddrA = FileAddrA
			tmpFileOffsA = FileOffsA
			tmpFileLenA = FileLenA
			tmpFileIOA = FileIOA

			SortBundle()

			'Restor arrays and lists

			Prgs = tmpPrgs.ToList
			FileNameA = tmpFileNameA
			FileAddrA = tmpFileAddrA
			FileOffsA = tmpFileOffsA
			FileLenA = tmpFileLenA
			FileIOA = tmpFileIOA

			ReDim Buffer(255)

			'Find the previous bundle's variables
TryAgain:
			BytePtr = BundleBytePtrA(CurrentBundle - PrevCP)
			BitPtr = BundleBitPtrA(CurrentBundle - PrevCP)
			BitPos = BundleBitPosA(CurrentBundle - PrevCP)
			PrevBC = BundleSizeA(CurrentBundle - PrevCP)

			If BytePtr = 0 Then
				If CurrentBundle - PrevCP > 0 Then
					PrevCP += 1
					GoTo TryAgain
				Else
					ResetBuffer()
				End If
			End If
		End If

		'Will need to finish the previous bundle here!!!
		If (BufferCnt = 0) And (BytePtr = 255) Then
			NewBlock = SetNewBlock          'SetNewBlock is true at closing the previous bundle, so first it just sets NewBlock2
			SetNewBlock = False             'And NewBlock will fire at the desired bundle
		Else
			Dim ThisPartIO As Integer = If(FileIOA.Count > 0, CheckNextIO(FileAddrA(0), FileLenA(0), FileIOA(0)), 0)
			CloseBundle(ThisPartIO, False, True)
		End If

		If BufferCnt = 1 Then
			'Closing the previous bundle resulted in an additional block, update previous bundle info
			BufferCnt = 0
			BundleBytePtrA(CurrentBundle - PrevCP) = BytePtr
			BundleBitPtrA(CurrentBundle - PrevCP) = BitPtr
			BundleBitPosA(CurrentBundle - PrevCP) = BitPos
			BundleSizeA(CurrentBundle - PrevCP) = PrevBC + 1
			PrevPartChgd = True 'We will add an additional block to the disk size
		End If

		CompressBundle(True)

		If CurrentBundle + 1 < PDiskNoA.Count Then
			If PDiskNoA(CurrentBundle) = PDiskNoA(CurrentBundle + 1) Then
				SetNewBlock = PNewBlockA(CurrentBundle + 1)
			End If
		Else
			SetNewBlock = False
		End If

		If SetNewBlock = True Then
			BufferCnt += 1
		End If

		If CurrentBundle > BundleBytePtrA.Count Then
			ReDim Preserve BundleBytePtrA(CurrentBundle), BundleBitPtrA(CurrentBundle), BundleBitPosA(CurrentBundle), BundleSizeA(CurrentBundle), BundleOrigSizeA(CurrentBundle)
		End If

		BundleBytePtrA(CurrentBundle) = BytePtr
		BundleBitPtrA(CurrentBundle) = BitPtr
		BundleBitPosA(CurrentBundle) = BitPos

		If Strings.Left(PartN.Name, 1) = "P" Then
			BufferCnt += 1
		End If

		BundleSizeA(CurrentBundle) = BufferCnt
		BundleOrigSizeA(CurrentBundle) = UncompBundleSize

		CalcPartSize = BufferCnt + If(PrevPartChgd = True, 1, 0)

		Exit Function
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Sub CalcDiskSizeWithForm(ParentNode As TreeNode, Optional PartNodeIndex As Integer = -1)

		Dim Frm As New FrmDisk
		Frm.Show(Me)

		StartUpdate()

		CalcDiskSize(ParentNode, PartNodeIndex)

		GoTo Done

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
		FinishUpdate()
		Frm.Close()

	End Sub

	Private Sub CalcDiskSize(ParentNode As TreeNode, Optional PartNodeIndex As Integer = -1)
		On Error GoTo Err

		'THIS WILL CALCULATE EVERY DISK'S AND Bundle'S SIZE IN THE SCRIPT AND UPDATE EVERY DISK AND Bundle NODE
		'ALWAYS CALL IT WITH BASENODE
		If ParentNode Is Nothing Then Exit Sub
		'If PartNode Is Nothing Then Exit Sub

		If ChkSize.Checked = False Then
			'No recalc, only renumber
			ReNumberEntries(BaseNode, PartNodeIndex)
			Exit Sub
		End If

		Dim StartIndex As Integer = 0
		If ParentNode.Name = BaseNode.Name Then
			ReNumberEntries(BaseNode, PartNodeIndex)

			'Packer = My.Settings.DefaultPacker
			ReDim DiskNodeA(DC), DiskSizeA(DC)
			'Default case - we calculate the whole script
			CurrentDisk = 0
			If PartNodeIndex > -1 Then
				BundleNode = ParentNode.Nodes(PartNodeIndex)
				'Calculate from selected node
				For I As Integer = PartNodeIndex To 0 Step -1
					'Find selected node's disk node and start calculations from this disk node
					If Int(BundleNode.Parent.Nodes(I).Tag / &H10000) = 1 Then
						DiskNode = BundleNode.Parent.Nodes(I)
						CurrentDisk = (DiskNode.Tag And &HFFFF)
						DiskNodeA(CurrentDisk) = DiskNode
						StartIndex = I + 1    'Exlude disk from calculation below
						'For J As Integer = 0 To DiskNode.Nodes.Count - 1
						'If InStr(DiskNode.Nodes(J).Name, sPacker) <> 0 Then
						'Select Case Strings.Right(LCase(DiskNode.Nodes(J).Text), 6)
						'Case "better"
						'Packer = 2
						'Case Else
						'Packer = 1
						'End Select
						'Exit For
						'End If
						'Next
						Exit For
					End If
				Next
			End If
		End If

		For I As Integer = StartIndex To ParentNode.Nodes.Count - 1
			Select Case Int(ParentNode.Nodes(I).Tag / &H10000)
				Case 1
					'Disk
					If CurrentDisk > 0 Then
						DiskNodeA(CurrentDisk).Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk).ToString + " block" +
							If(DiskSizeA(CurrentDisk) = 1, "", "s") + " used, " + (MaxDiskSize - DiskSizeA(CurrentDisk)).ToString + " block" +
							If((MaxDiskSize - DiskSizeA(CurrentDisk)) = 1, "", "s") + " free]"
					End If
					DiskNode = ParentNode.Nodes(I)
					CurrentDisk = (DiskNode.Tag And &HFFFF)
					DiskNodeA(CurrentDisk) = DiskNode
					DiskSizeA(CurrentDisk) = 0
					'For J As Integer = 0 To DiskNode.Nodes.Count - 1
					'If InStr(DiskNode.Nodes(J).Name, sPacker) <> 0 Then
					'Select Case Strings.Right(LCase(DiskNode.Nodes(J).Text), 6)
					'Case "better"
					'Packer = 2
					'Case Else
					'Packer = 1
					'End Select
					'Exit For
					'End If
					'Next
				Case 2
					'Bundle
					BundleNode = ParentNode.Nodes(I)
					CurrentBundle = (BundleNode.Tag And &HFFFF)
					Dim CPS As Integer = CalcPartSize(ParentNode.Nodes(I))
					If DiskSizeA(CurrentDisk) + CPS > MaxDiskSize Then
						BundleNode.Text = "[Bundle " + CurrentBundle.ToString + "]"
						MsgBox("The size of this disk exceeds 664 blocks!", vbOKOnly + vbCritical, "Disk is full!")
						Exit For
					End If
					DiskSizeA(CurrentDisk) += CPS
				Case 3
					'Script
					CalcDiskSize(ParentNode.Nodes(I))
			End Select
		Next

		If ParentNode.Name = BaseNode.Name Then
			If CurrentDisk > 0 Then
				DiskNodeA(CurrentDisk).Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk).ToString + " block" +
							If(DiskSizeA(CurrentDisk) = 1, "", "s") + " used, " + (MaxDiskSize - DiskSizeA(CurrentDisk)).ToString + " block" +
							If((MaxDiskSize - DiskSizeA(DC)) = 1, "", "s") + " free]"
			End If
			UpdatePartNames(BaseNode)
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub UpdatePartNames(ParentNode As TreeNode)
		On Error GoTo Err

		For I As Integer = 0 To ParentNode.Nodes.Count - 1
			Select Case Int(ParentNode.Nodes(I).Tag / &H10000)
				Case 1
					'Disk
				Case 2
					'Bundle
					BundleNode = ParentNode.Nodes(I)
					CurrentBundle = (BundleNode.Tag And &HFFFF)
					BufferCnt = BundleSizeA(CurrentBundle)
					UncompBundleSize = BundleOrigSizeA(CurrentBundle)

					BundleNode.Text = "[Bundle " + Strings.Right(BundleNode.Name, Len(BundleNode.Name) - 1) + ": " + BufferCnt.ToString +
					" block" + If(BufferCnt = 1, "", "s") + " packed from " + UncompBundleSize.ToString + " block" + If(UncompBundleSize = 1, "", "s") +
					" unpacked, " + (Int(10000 * BufferCnt / UncompBundleSize) / 100).ToString + "% of unpacked size]"
				Case 3
					'Script
					UpdatePartNames(ParentNode.Nodes(I))
			End Select
		Next

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub TxtEdit_MouseHover(sender As Object, e As EventArgs) Handles txtEdit.MouseHover
		On Error GoTo Err

		ShowTT()

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ShowTT()
		On Error GoTo Err

		TT.Hide(txtEdit)

		If ChkToolTips.Checked = False Then Exit Sub

		If Loading = True Then Exit Sub

		Dim TTT As String = ""

		With TT
			.ToolTipIcon = ToolTipIcon.Info
			.UseFading = True
			.InitialDelay = 1000
			'.AutomaticDelay = 1000
			.AutoPopDelay = 5000
			.ReshowDelay = 1000
			Select Case txtEdit.Tag
				Case sDiskHeader
					.ToolTipTitle = "Editing the Disk's Header"
					TTT = "Type in the disk's header (max. 16 characters). Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sDiskID
					.ToolTipTitle = "Editing the Disk's ID"
					TTT = "Type in the disk's ID (max. 5 characters). Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sDemoName
					.ToolTipTitle = "Editing the Demo's Name"
					TTT = "Type in the demo's name  (max. 16 characters) which will be shown as the first PRG's name in the directory." +
							vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sDemoStart + "$"
					.ToolTipTitle = "Editing the Start Address of the Demo"
					TTT = "Type in the demo's entry point to be used when the disk is started from Basic." +
							 vbNewLine + "If this field is left empty, Sparkle will use the first file's load address as entry point." +
							 vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sFileAddr
					.ToolTipTitle = "Editing the file segment's Load Address"
					TTT = "Type in the hex load address of this data segment." +
							vbNewLine + "If this field is left empty, Sparkle will reset it to its default value." +
							vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sFileOffs
					.ToolTipTitle = "Editing the file segment's Offset"
					TTT = "Type in the hex offset of this data segment (first byte to be loaded)." +
							vbNewLine + "If this field is left empty, Sparkle will reset it to its default value." +
							vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sFileLen
					.ToolTipTitle = "Editing the file segment's Length"
					TTT = "Type in the hex length of this data segment." +
							vbNewLine + "If this field is left empty, Sparkle will use (file length-offset) as length." +
							vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sZP + "$"
					.ToolTipTitle = "Editing the Zeropage Usage of the Loader"
					TTT = "Type in the first of the two adjacent zeropage addresses you want the loader to use." +
							 vbNewLine + "If this field is left empty, Sparkle will use $02-$03 as default." +
							 vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case sLoop
					.ToolTipTitle = "Editing where the demo should loop after loading the last disk"
					TTT = "Type in a decimal number between 0 and 255." +
							vbNewLine + "The default value of 0 will terminate the demo without looping." +
							vbNewLine + "A number between 1 and 255 will specify the disk the demo will loop to." +
							vbNewLine + "Press <Enter> or <Tab> to save changes, or <Escape> to cancel editing."
				Case Else
					Exit Select
			End Select

			If TTT <> "" Then
				.Show(TTT, txtEdit)
			End If
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub
	Private Sub ToggleFileNodes(N As TreeNode)
		On Error GoTo Err

		StartUpdate()

		If N.Name = BaseNode.Name Then
			Loading = True
		End If

		If ChkExpand.Checked = False Then
			For I As Integer = 0 To N.Nodes.Count - 1
				N.Nodes(I).Collapse()
				Select Case Int(N.Nodes(I).Tag / &H10000)
					Case 0
					Case 1
					Case 2
						For J As Integer = 1 To N.Nodes(I).Nodes.Count - 1
							N.Nodes(I).Nodes(J).Collapse()
						Next
					Case 3
						ToggleFileNodes(N.Nodes(I))
				End Select
			Next
			'N.Nodes(N.Nodes.Count - 1).Expand()
		Else
			For I As Integer = 0 To N.Nodes.Count - 1
				N.Nodes(I).Expand()
				Select Case Int(N.Nodes(I).Tag / &H10000)
					Case 0
					Case 1
					Case 2
						For J As Integer = 1 To N.Nodes(I).Nodes.Count - 1
							N.Nodes(I).Nodes(J).Expand()
						Next
					Case 3
						ToggleFileNodes(N.Nodes(I))
				End Select
			Next
		End If

		If N.Name = BaseNode.Name Then
			Loading = False
		End If

		FinishUpdate()

		TV.Focus()

		If TV.SelectedNode IsNot Nothing Then
			TV.SelectedNode.EnsureVisible()
		End If

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnNew_MouseEnter(sender As Object, e As EventArgs) Handles BtnNew.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press Ctrl+N to start a new script."
		ShowToolTip("New Script", TTT, BtnNew)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnLoad_MouseEnter(sender As Object, e As EventArgs) Handles BtnLoad.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press Ctrl+L to load an existing script."
		ShowToolTip("Load Script", TTT, BtnLoad)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnSave_MouseEnter(sender As Object, e As EventArgs) Handles BtnSave.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press Ctrl+S to save your script."
		ShowToolTip("Save Script", TTT, BtnSave)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnEntryDown_MouseEnter(sender As Object, e As EventArgs) Handles BtnEntryDown.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press Ctrl+PgDn to move this entry and its content down in your script."
		ShowToolTip("Move Entry Down", TTT, BtnEntryDown)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnEntryUp_MouseEnter(sender As Object, e As EventArgs) Handles BtnEntryUp.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press Ctrl+PgUp to move this entry and its content up in your script."
		ShowToolTip("Move Entry Up", TTT, BtnEntryUp)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ChkSize_MouseEnter(sender As Object, e As EventArgs) Handles ChkSize.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Check or press Ctrl+A to autocalculate disk and bundle sizes during script editing."
		ShowToolTip("Autocalculate Disk and Bundle Sizes", TTT, ChkSize)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ChkExpand_MouseEnter(sender As Object, e As EventArgs) Handles ChkExpand.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Check or press Ctrl+D to show entry details."
		ShowToolTip("Show Details", TTT, ChkExpand)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ChkToolTips_MouseEnter(sender As Object, e As EventArgs) Handles ChkToolTips.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Check or press Ctrl+T to show tooltips."
		ShowToolTip("Show TooTips", TTT, ChkToolTips)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub ShowToolTip(TTTitle As String, TTText As String, Owner As Control)
		On Error GoTo Err

		If ChkToolTips.Checked = False Then Exit Sub

		With TT
			.ToolTipIcon = ToolTipIcon.Info
			.UseFading = True
			.InitialDelay = 1000        '=AutomaticDelay
			'.AutomaticDelay = 1000
			.AutoPopDelay = 5000        '=AD x 10
			.ReshowDelay = 1000         '=AD/5
			.ToolTipTitle = TTTitle
			.Show(TTText, Owner)
		End With

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub OptFullPaths_MouseEnter(sender As Object, e As EventArgs) Handles OptFullPaths.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Check or press Ctrl+P to save your scripts with full file paths."
		ShowToolTip("Use Full Paths", TTT, OptFullPaths)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub OptRelativePaths_MouseEnter(sender As Object, e As EventArgs) Handles OptRelativePaths.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Check or press Ctrl+R to save your scripts with file paths relative to the script's path."
		ShowToolTip("Use Relative Paths", TTT, OptRelativePaths)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnOK_MouseEnter(sender As Object, e As EventArgs) Handles BtnOK.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press F5 to close the editor and build your demo."
		ShowToolTip("Close & Build Demo", TTT, BtnOK)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub BtnCancel_MouseEnter(sender As Object, e As EventArgs) Handles BtnCancel.MouseEnter
		On Error GoTo Err

		Dim TTT As String = "Click or press Alt+F4 to closed the editor without building your demo."
		ShowToolTip("Close without Building Demo", TTT, BtnCancel)

		Exit Sub
Err:
		ErrCode = Err.Number
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

End Class