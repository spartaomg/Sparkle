Imports System.ComponentModel
Public Class FrmSE

    'Private Const WM_USER As Integer = &H400
    'Private Const EN_HSCROLL As Integer = &H601
    'Private Const EN_VSCROLL As Integer = &H602
    'Private Const EN_CHANGE As Integer = &H300
    'Private Const EN_UPDATE As Integer = &H400
    'Private Const EM_SETEVENTMASK = (WM_USER + 69)
    'Private Const ENM_SCROLL As Integer = &H4
    'Private Const EM_GETSCROLLPOS As Integer = (WM_USER + 221)
    'Private Const EM_SETSCROLLPOS As Integer = (WM_USER + 222)
    'Private Const WM_COMMAND As Integer = &H111
    'Private Const WM_LBUTTONDOWN As Integer = &H201
    'Private Const WM_CLOSE As Integer = &H10
    'Private Const WM_QUERYENDSESSION As Integer = &H11
    'Private Const SC_CLOSE As Integer = &HF060&
    'Private Const WM_SYSCOMMAND As Integer = &H112
    'Private Const WM_KEYDOWN As Integer = &H100
    'Private Const WM_DESTROY As Integer = &H2
    'Private Const SC_MAXIMIZE As Integer = &HF030
    'Private Const SC_MINIMIZE As Integer = &HF020
    'Private Const SC_RESTORE As Integer = &HF120
    Private Const WM_VSCROLL As Integer = &H115
    Private Const WM_HSCROLL As Integer = &H114
    Private Const WM_MOUSEWHEEL As Integer = &H20A

    Private ReadOnly DefaultCol As Color = Color.RosyBrown
    Private ReadOnly ManualCol As Color = Color.SaddleBrown

    Private WithEvents SCC As SubClassCtrl.SubClassing
    Private SelNode As TreeNode

    Private FileType As Byte
    Private KeyEnter As Boolean
    Private DC, PC, FC As Integer
    Private FileSize As Integer
    Private FilePath As String
    Private KeyID, PartID, FileID As String
    Private NodeType As Byte
    Private FAddr, FOffs, FLen, PLen As Integer
    Private FAS, FOS, FLS As String
    Private Loading As Boolean = False
    Private txtBuffer As String = ""
    Private LMD As Date = Date.Now
    Private Dbl As Boolean = False
    'Private DefaultParams As Boolean = True
    Private DFA, DFO, DFL As Boolean

    Private DFAS, DFOS, DFLS As String
    Private DFAN, DFON, DFLN As Integer

    Private ReadOnly sDiskPath As String = "Disk Path: "
    Private ReadOnly sDiskHeader As String = "Disk Header: "
    Private ReadOnly sDiskID As String = "Disk ID: "
    Private ReadOnly sDemoName As String = "Demo Name: "
    Private ReadOnly sDemoStart As String = "Demo Start: "
    Private ReadOnly sAddDisk As String = "AddDisk"
    Private ReadOnly sAddPart As String = "AddPart"
    Private ReadOnly sAddFile As String = "AddFile"
    Private ReadOnly sFileSize As String = "Original File Size: "
    Private ReadOnly sFileUIO As String = "Load under I/O: "
    Private ReadOnly sFileAddr As String = "Load Address: $"
    Private ReadOnly sFileOffs As String = "File Offset:  $"
    Private ReadOnly sFileLen As String = "File Length:  $"
    Private ReadOnly sDirArt As String = "DirArt: "
    Private ReadOnly sZP As String = "Zeropage: "
	Private ReadOnly sPacker As String = "Packer: "

	Private ReadOnly TT As New ToolTip

    Private ReadOnly tDiskPath As String = "Double click or press <Enter> to specify where your demo disk will be saved in D64 format."
    Private ReadOnly tDiskHeader As String = "Double click or press <Enter> to edit the disk's header."
    Private ReadOnly tDiskID As String = "Double click or press <Enter> to edit the disk's ID."
    Private ReadOnly tDemoName As String = "Double click or press <Enter> to edit the name of the first PRG in the directory."
    Private ReadOnly tDemoStart As String = "Double click or press <Enter> to edit the start address (entry point) of the demo."
    Private ReadOnly tAddDisk As String = "Double click or press <Enter> to add a new disk structure to the script."
    Private ReadOnly tAddPart As String = "Double click or press <Enter> to add a new part to this disk."
    Private ReadOnly tAddFile As String = "Double click or press <Enter> to add a new file to this demo part."
    Private ReadOnly tFileSize As String = "Original size of the selected file segment in blocks."
    Private ReadOnly tFileAddr As String = "Double click or press <Enter> to edit the file segment's load address."
    Private ReadOnly tFileOffs As String = "Double click or press <Enter> to edit the file segment's offset."
    Private ReadOnly tFileLen As String = "Double click or press <Enter> to edit the file segment's length."
    Private ReadOnly tLoadUIO As String = "Double click or press <Enter> to change the file segment's I/O status, if applicable."
    Private ReadOnly tDirArt As String = "Double click or press <Enter> to add a DirArt file to the demo's directory." + vbNewLine +
                "Press the <Delete> key to delete the current DirArt file."
    Private ReadOnly tDisk As String = "Press <Delete> to delete this disk with all its content."
    Private ReadOnly tPart As String = "Files and file segments in this part will be loaded during a single loader call." + vbNewLine +
                "Press <Delete> to delete this part from this disk with all its content."
    Private ReadOnly tFile As String = "Double click or press <Enter> to change this file." + vbNewLine +
                "Press <Delete> to delete this file from this part."
    Private ReadOnly tZP As String = "Double click or press <Enter> to edit the loader's zeropage usage."
	Private ReadOnly tPacker As String = "Double click or press <Enter> to change the loader's packer selection." + vbNewLine +
				"The 'faster' option results in a faster but less effective compression and somewhat faster loading." + vbNewLine +
				"The 'better' option results in a slower but more effective compression and somewhat slower loading."

	Private Sub FrmSE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo Err

		If My.Settings.EditorWindowMax = True Then
            WindowState = FormWindowState.Maximized
        Else
            WindowState = FormWindowState.Normal
            Width = My.Settings.EditorWidth
            Height = My.Settings.EditorHeight
        End If

		If My.Settings.DefaultPacker = 1 Then
			OptFaster.Checked = True
		Else
			OptBetter.Checked = True
		End If

		With TT
            .ToolTipIcon = ToolTipIcon.Info
            .UseFading = True
            .InitialDelay = 2000
            .AutomaticDelay = 2000
            .AutoPopDelay = 1000
            .ReshowDelay = 1000
        End With

        Refresh()

        bBuildDisk = False

        tv.AllowDrop = True

        chkExpand.Checked = My.Settings.ShowFileDetails
        chkToolTips.Checked = My.Settings.ShowToolTips

        DC = 0
        PC = 0
        FC = 0

        CurrentDisk = DC
        CurrentPart = PC
        CurrentFile = FC

        ReDim PartSizeA(PC), PartByteCntA(PC), PartBitCntA(PC), PartBitPosA(PC)

        PartByteCntA(PC) = 254
        PartBitCntA(PC) = 0
        PartBitPosA(PC) = 15

        DiskCnt = DC
        ReDim DiskSizeA(DiskCnt)

        AddNewDiskNode()

        If (Script <> "") And (Script <> ScriptHeader + vbNewLine + vbNewLine) Then
            ConvertScriptToNodes()
            Tv_GotFocus(sender, e)
        Else
            tv.SelectedNode = tv.Nodes(sAddDisk)
            BlankDiskStructure()
        End If

        tssLabel.Text = If(ScriptName <> "", "Script: " + ScriptName, "Script: (new script)")

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles tv.AfterSelect
        On Error GoTo Err

        NodeSelect()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_GotFocus(sender As Object, e As EventArgs) Handles tv.GotFocus
        On Error GoTo Err

        NodeSelect()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub NodeSelect()
        On Error GoTo Err

        If tv.SelectedNode Is Nothing Then Exit Sub

        SelNode = tv.SelectedNode

        CurrentDisk = Int(SelNode.Tag / &H1000000)
        CurrentPart = Int((SelNode.Tag And &HFFF000) / &H1000)
        CurrentFile = SelNode.Tag And &HFFF

        If (tv.Enabled = False) Or (Loading) Then Exit Sub

        Select Case SelNode.ForeColor
            Case Color.DarkRed      'Disk node
                NodeType = 1
            Case Color.DarkMagenta  'Part node
                NodeType = 2
            Case Color.Black        'File node
                NodeType = 3
            Case Else
                If Strings.Left(SelNode.Text, 8) = "DirArt: " Then
                    NodeType = 4    'DirArt node
                Else
                    NodeType = 0    'All other nodes
                End If
        End Select

        If NodeType = 3 Then
            BtnFileUp.Enabled = SelNode.Index > 0
            BtnFileDown.Enabled = SelNode.Index < SelNode.Parent.Nodes.Count - 2
        Else
            BtnFileDown.Enabled = False
            BtnFileUp.Enabled = False
        End If

        If NodeType = 2 Then
            Dim NI As Integer
            For I As Integer = 0 To SelNode.Parent.Nodes.Count - 1
                If InStr(SelNode.Parent.Nodes(I).Text, "[Part") <> 0 Then
                    NI = I
                    Exit For
                End If
            Next
            BtnPartUp.Enabled = SelNode.Index > NI
            BtnPartDown.Enabled = SelNode.Index < SelNode.Parent.Nodes.Count - 2
        Else
            BtnPartDown.Enabled = False
            BtnPartUp.Enabled = False
        End If

        If CurrentDisk > 0 Then
            TssDisk.Text = "Disk " + CurrentDisk.ToString + ": " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s free", " free")
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub DeleteNode(N As TreeNode)
        On Error GoTo Err

        Dim P As TreeNode

        Select Case NodeType
            Case 0
                'MsgBox("This entry cannot be deleted!", vbInformation + vbOKOnly)
            Case 1  'Disk
                If MsgBox("Are you sure you want to delete Disk " + CurrentDisk.ToString + " and all its content?" + vbNewLine + vbNewLine + N.Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    N.Remove()
                End If
            Case 2  'Part
                If MsgBox("Are you sure you want to delete Part " + CurrentPart.ToString + " and all its content?" + vbNewLine + vbNewLine + N.Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    P = N.Parent.Nodes(N.Index + 1)
                    N.Remove()
                    tv.Refresh()
                    CalcPartSize(P)
                End If
            Case 3  'File
                If MsgBox("Are you sure you want to delete the following file entry?" + vbNewLine + vbNewLine + N.Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    P = N.Parent
                    N.Remove()
                    tv.Refresh()
                    CalcPartSize(P)
                End If
            Case 4  'DirArt
                If N.Text <> sDirArt Then
                    If MsgBox("Are you sure you want to delete the following DirArt file?" + vbNewLine + vbNewLine + Strings.Right(N.Text, N.Text.Length - sDirArt.Length), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        N.Text = sDirArt
                    End If
                End If
        End Select

        NodeSelect()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_KeyDown(sender As Object, e As KeyEventArgs) Handles tv.KeyDown
        On Error GoTo Err

        Dim S As String
        Dim N As TreeNode = tv.SelectedNode

        If e.KeyCode = Keys.Delete Then
            KeyEnter = True
            e.SuppressKeyPress = True
            DeleteNode(N)
            Exit Sub
        ElseIf e.KeyCode <> Keys.Enter Then
            KeyEnter = False
            Exit Sub
        End If

        KeyEnter = True
        e.SuppressKeyPress = True

        Select Case Strings.Left(N.Text, 13)
            Case "[Add new disk"   'Add new disk
                txtEdit.Visible = False
                If MsgBox("Do you want to add a new disk?", vbYesNo + vbQuestion, "Add new disk?") = vbYes Then BlankDiskStructure()
                Exit Sub
            Case "[Add new part"   'Add new part to current disk
                txtEdit.Visible = False
                UpdateNewPartNode()
                Exit Sub
            Case "[Add new file"   'Add new file to current part
                txtEdit.Visible = False
                AddNewFile()
                'ToggleFileNodesC
                Exit Sub
            Case Else
                Exit Select
        End Select

        S = Strings.Left(N.Name, InStr(N.Name, ":") + 1)

        If Strings.Right(N.Name, 3) = ":FS" Then        'Is this a File Size node?
            txtEdit.Visible = False
            Exit Sub
        ElseIf Strings.Right(N.Name, 5) = ":FUIO" Then
            txtEdit.Visible = False
            SwapIOStatus()
            Exit Sub
        ElseIf Strings.Right(N.Name, 3) = ":FA" Then
            'Load Address
            S = sFileAddr
FileData:
            With txtEdit
                .Tag = S
                .Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
                N.Text = S
                .Top = tv.Top + N.Bounds.Top + 3
                .Left = tv.Left + N.Bounds.Left + N.Bounds.Width
                .Width = TextRenderer.MeasureText(.Text, N.NodeFont).Width
                .ForeColor = N.ForeColor
                .MaxLength = 4
                .Visible = True
            End With
        ElseIf Strings.Right(N.Name, 3) = ":FO" Then
            'File Offset
            S = sFileOffs
            GoTo FileData
        ElseIf Strings.Right(N.Name, 3) = ":FL" Then
            'File Length
            S = sFileLen
            GoTo FileData
        ElseIf Strings.InStr(N.Name, ":F") > 0 Then   'Is this a File node?
            txtEdit.Visible = False
            UpdateFileNode()
            Exit Sub
        Else
            'Any other nodes
            With txtEdit
                .MaxLength = 32767
                Select Case S
                    Case sDiskPath
                        FilePath = Strings.Right(N.Text, Len(N.Text) - Len(S))
                        .Tag = S
                        .Visible = False    'True
                        UpdateDiskPath()
                        Exit Sub
                    Case sDiskHeader
                        .Tag = S
                        .Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
                        N.Text = S
                        .Left = tv.Left + N.Bounds.Left + N.Bounds.Width
                        .Width = tv.Left + tv.Width - .Left - 2 - 17    'Subtract border and scrollbar widths
                        .MaxLength = 16
                        .Visible = True
                    Case sDiskID
                        .Tag = S
                        .Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
                        N.Text = S
                        .Left = tv.Left + N.Bounds.Left + N.Bounds.Width
                        .Width = tv.Left + tv.Width - .Left - 2 - 17    'Subtract border and scrollbar widths
                        .MaxLength = 5
                        .Visible = True
                    Case sDemoName
                        .Tag = S
                        .Text = Strings.Right(N.Text, Len(N.Text) - Len(S))
                        N.Text = S
                        .Left = tv.Left + N.Bounds.Left + N.Bounds.Width
                        .Width = tv.Left + tv.Width - .Left - 2 - 17    'Subtract border and scrollbar widths
                        .MaxLength = 16
                        .Visible = True
                    Case sDemoStart
                        .Text = Strings.Right(N.Text, Len(N.Text) - Len(S) - 1)
                        .Width = TextRenderer.MeasureText("0000", N.NodeFont).Width
                        .Tag = S + "$"
                        N.Text = .Tag
                        .Left = tv.Left + N.Bounds.Left + N.Bounds.Width
                        .MaxLength = 4
                        .Visible = True
                    Case sDirArt
                        FilePath = Strings.Right(N.Text, Len(N.Text) - Len(S))
                        .Tag = S
                        .Visible = False    'True
                        UpdateDirArtPath()
                        Exit Sub
                    Case sZP
                        .Text = Strings.Right(N.Text, Len(N.Text) - Len(S) - 1)
                        .Width = TextRenderer.MeasureText("00", N.NodeFont).Width
                        .Tag = S + "$"
                        N.Text = .Tag
                        .Left = tv.Left + N.Bounds.Left + N.Bounds.Width
                        .MaxLength = 2
						.Visible = True
					Case sPacker
						.Visible = False
						Select Case LCase(Strings.Right(N.Text, 6))
							Case "faster"
								N.Text = sPacker + "better"
							Case "better"
								N.Text = sPacker + "faster"
						End Select
					Case Else
						.Visible = False
                End Select
                .Top = tv.Top + N.Bounds.Top + 3
                .ForeColor = N.ForeColor
            End With
        End If

        If txtEdit.Visible = True Then
            txtEdit.Refresh()
            txtEdit.Focus()
            txtEdit.SelectAll()
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tv.KeyPress
        On Error GoTo Err

        If KeyEnter = True Then e.Handled = True

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tv.NodeMouseDoubleClick
        On Error GoTo Err

        Dim k As New KeyEventArgs(Keys.Enter)
        Tv_KeyDown(sender, k)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub TxtEdit_KeyDown(sender As Object, e As KeyEventArgs) Handles txtEdit.KeyDown
        On Error GoTo Err

        Select Case e.KeyCode
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down, Keys.Back, Keys.Delete
            Case Keys.Enter
                e.SuppressKeyPress = True
                e.Handled = True
                tv.Focus()
            Case Keys.Escape
                txtEdit.Text = txtBuffer
                e.SuppressKeyPress = True
                e.Handled = True
                tv.Focus()
            Case Else
                If txtEdit.MaxLength = 4 Or txtEdit.MaxLength = 2 Then
                    Select Case e.KeyCode
                        Case Keys.A, Keys.B, Keys.C, Keys.D, Keys.E, Keys.F
                        Case Keys.D0, Keys.D1, Keys.D2, Keys.D3, Keys.D4, Keys.D5, Keys.D6, Keys.D7, Keys.D8, Keys.D9
                        Case Else
                            e.SuppressKeyPress = True
                            e.Handled = True
                    End Select
                End If
        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub TxtEdit_GotFocus(sender As Object, e As EventArgs) Handles txtEdit.GotFocus
        On Error GoTo Err

        SCC = New SubClassCtrl.SubClassing(tv.Handle) With {
        .SubClass = True
        }

        txtBuffer = txtEdit.Text

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub TxtEdit_LostFocus(sender As Object, e As EventArgs) Handles txtEdit.LostFocus
        On Error GoTo Err

        SCC.ReleaseHandle()

        tv.Focus()

        CorrectTextLength()

        SelNode.Text = txtEdit.Tag + txtEdit.Text

        tv.Invalidate(SelNode.Bounds)   'This is to repaint selnode

        txtEdit.Visible = False

        Select Case Strings.Right(SelNode.Name, 3)
            Case ":FA", ":FO", ":FL"
                ChangeFileParameters(SelNode.Parent, SelNode.Index)
                Exit Sub
        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub CorrectTextLength()
        On Error GoTo Err

        'Corrects the length of txtEdit to 2 or 4 characters depending on MaxLength
        'Eliminates invalid values

        Select Case txtEdit.MaxLength
            Case 2  'ZP value
                If txtEdit.Text.Length < 2 Then
                    txtEdit.Text = Strings.Left("02", 2 - txtEdit.Text.Length) + txtEdit.Text
                End If

                'Invalid values: 00, 01, ff
                If txtEdit.Text = "00" Or txtEdit.Text = "01" Then
                    txtEdit.Text = "02"
                ElseIf txtEdit.Text = "ff" Then
                    txtEdit.Text = "fe"
                End If
            Case 4  'Address values
                If (txtEdit.Text.Length < 4) And (txtEdit.Text.Length > 0) Then
                    txtEdit.Text = Strings.Left("0000", 4 - txtEdit.Text.Length) + txtEdit.Text
                End If
            Case Else
                Exit Sub
        End Select

        txtEdit.Text = LCase(txtEdit.Text)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ChangeFileParameters(FileNode As TreeNode, NodeIndex As Integer)
        On Error GoTo Err

        If Loading = False Then tv.BeginUpdate()

        'If node is edited, then A/O/L=txtEdit, otherwise A/O/L = 4 rightmost chars of node text
        Dim A As String = IIf(NodeIndex = 0, txtEdit.Text, Strings.Right(FileNode.Nodes(0).Text, FileNode.Nodes(0).Text.Length - sFileAddr.Length))
        Dim O As String = IIf(NodeIndex = 1, txtEdit.Text, Strings.Right(FileNode.Nodes(1).Text, FileNode.Nodes(1).Text.Length - sFileOffs.Length))
        Dim L As String = IIf(NodeIndex = 2, txtEdit.Text, Strings.Right(FileNode.Nodes(2).Text, FileNode.Nodes(2).Text.Length - sFileLen.Length))

        GetDefaultFileParameters(FileNode, A, O, L)

        If txtEdit.Text = "" Then
            ResetFileParameters(FileNode, NodeIndex)
        ElseIf txtEdit.Text <> txtBuffer Then
            Select Case NodeIndex
                Case 0
                    'Load Address has changed, check if Offset is default and update it as needed
                    If FileNode.Nodes(1).ForeColor = DefaultCol Then FileNode.Nodes(1).Text = sFileOffs + DFOS
                    GoTo Node2
                Case 1
                    'Load Address and/or Offset have changed, check if length is default and update it as needed
Node2:              If FileNode.Nodes(2).ForeColor = DefaultCol Then FileNode.Nodes(2).Text = sFileLen + DFLS
                Case 2
                    'Nothing here...
            End Select
        End If

        ValidateFileParameters(FileNode)    'This will make sure parameters are within limits
        'It will also update file size and part size(s)

        CheckFileParameterColors(FileNode)  'Check if parameters are default or not

        GoTo Done
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:
        If Loading = False Then tv.EndUpdate()

    End Sub

    Private Sub CheckFileParameterColors(FileNode As TreeNode)
        On Error GoTo Err

        If Loading = False Then tv.BeginUpdate()

        If Strings.Right(FileNode.Nodes(0).Text, 4) = DFAS Then
            DFA = True
        Else
            DFA = False
        End If

        If Strings.Right(FileNode.Nodes(1).Text, 4) = DFOS Then
            DFO = True
            If DFOS = "0000" Then
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
        FileNode.Nodes(0).ForeColor = IIf(DFA = True, DefaultCol, ManualCol)
        FileNode.Nodes(1).ForeColor = IIf(DFO = True, DefaultCol, ManualCol)
        FileNode.Nodes(2).ForeColor = IIf(DFL = True, DefaultCol, ManualCol)

        GoTo Done
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:
        If Loading = False Then tv.EndUpdate()

    End Sub

    Private Sub ResetFileParameters(FileNode As TreeNode, NodeIndex As Integer)
        On Error GoTo Err

        If Loading = False Then tv.BeginUpdate()

        Select Case NodeIndex
            Case 0
                txtEdit.Text = DFAS
                With FileNode.Nodes(0)
                    .Text = sFileAddr + DFAS
                    .ForeColor = DefaultCol
                End With
                With FileNode.Nodes(1)
                    .Text = sFileOffs + DFOS
                    .ForeColor = DefaultCol
                End With
                With FileNode.Nodes(2)
                    .Text = sFileLen + DFLS
                    .ForeColor = DefaultCol
                End With

            Case 1
                txtEdit.Text = DFOS
                With FileNode.Nodes(1)
                    .Text = sFileOffs + DFOS
                    .ForeColor = DefaultCol
                End With
                With FileNode.Nodes(2)
                    .Text = sFileLen + DFLS
                    .ForeColor = DefaultCol
                End With
            Case 2
                txtEdit.Text = DFLS
                With FileNode.Nodes(2)
                    .Text = sFileLen + DFLS
                    .ForeColor = DefaultCol
                End With

            Case Else
                Exit Select
        End Select

        GoTo Done
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
        If Loading = False Then tv.EndUpdate()

    End Sub

    Private Sub GetDefaultFileParameters(FileNode As TreeNode, Optional FA As String = "", Optional FO As String = "", Optional FL As String = "")
        On Error GoTo Err

        If Loading = False Then tv.BeginUpdate()

        Dim P() As Byte = IO.File.ReadAllBytes(Replace(FileNode.Text, "*", ""))
        Dim Ext As String = LCase(Strings.Right(Replace(FileNode.Text, "*", ""), 4))

        PLen = P.Length

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

        'Calculate default parameter strings
        DFAS = ConvertNumberToHexString(DFAN Mod 256, Int(DFAN / 256))
        DFOS = ConvertNumberToHexString(DFON Mod 256, Int(DFON / 256))
        DFLS = ConvertNumberToHexString(DFLN Mod 256, Int(DFLN / 256))

        GoTo Done
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

Done:
        If Loading = False Then tv.EndUpdate()

    End Sub

    Private Sub ValidateFileParameters(FileNode As TreeNode)
        On Error GoTo Err

        NewFile = FileNode.Text

        FAddr = Convert.ToInt32(Strings.Right(FileNode.Nodes(0).Text, 4), 16)
        FOffs = Convert.ToInt32(Strings.Right(FileNode.Nodes(1).Text, 4), 16)
        FLen = Convert.ToInt32(Strings.Right(FileNode.Nodes(2).Text, 4), 16)

        If Loading = False Then tv.BeginUpdate()

        'Make sure Offset is within program length
        If FOffs > PLen - 1 Then
            FOffs = PLen - 1
            FileNode.Nodes(1).Text = sFileOffs + ConvertNumberToHexString(FOffs Mod 256, Int(FOffs / 256))
        End If

        'Check file length
        If (FLen = 0) Or (FOffs + FLen > PLen) Then
            FLen = PLen - FOffs
            DFLN = FLen
            DFLS = ConvertNumberToHexString(DFLN Mod 256, Int(DFLN / 256))
            FileNode.Nodes(2).Text = sFileLen + ConvertNumberToHexString(FLen Mod 256, Int(FLen / 256))
        End If

        'Make sure file is within memory
        If FAddr + FLen > &HFFFF Then
            FLen = &H10000 - FAddr
            FileNode.Nodes(2).Text = sFileLen + ConvertNumberToHexString(FLen Mod 256, Int(FLen / 256))
        End If

        FAS = ConvertNumberToHexString(FAddr Mod 256, Int(FAddr / 256))
        FOS = ConvertNumberToHexString(FOffs Mod 256, Int(FOffs / 256))
        FLS = ConvertNumberToHexString(FLen Mod 256, Int(FLen / 256))

        FileNode.Text = CalcFileSize(FileNode.Text, FAddr, FLen)
        FileNode.Nodes(4).Text = sFileSize + FileSize.ToString + " block" + IIf(FileSize = 1, "", "s")

        With FileNode
            If UnderIO() = False Then
                .Text = Replace(.Text, "*", "")
                .Nodes(3).Text = sFileUIO + "n/a"
                .Nodes(3).ForeColor = Color.MediumPurple
            ElseIf InStr(.Text, "*") <> 0 Then
                .Nodes(3).Text = sFileUIO + "yes"
                .Nodes(3).ForeColor = Color.Purple
            Else
                .Nodes(3).Text = sFileUIO + " no"
                .Nodes(3).ForeColor = Color.MediumPurple
            End If
        End With

        CalcPartSize(FileNode.Parent)

        GoTo Done
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
Done:
        If Loading = False Then tv.EndUpdate()

    End Sub

    Private Sub BlankDiskStructure()
        On Error GoTo Err

        tv.Enabled = False

        Dim N As TreeNode = tv.SelectedNode

        DC += 1

        If DiskSizeA.Count < DC Then
            ReDim Preserve DiskSizeA(DC - 1)
        End If

        CurrentDisk = DC

        tv.BeginUpdate()
        With N
            .Text = "[Disk " + DC.ToString + "]"
            .Name = "D" + DC.ToString
            .ForeColor = Color.DarkRed
            .Tag = DC * &H1000000
        End With

        Dim Fnt As New Font("Consolas", 10)

        AddNode(N, sDiskPath + DC.ToString, sDiskPath + "C:\demo.d64", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDiskHeader + DC.ToString, sDiskHeader + "demo disk " + Year(Now).ToString, N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDiskID + DC.ToString, sDiskID + "sprkl", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDemoName + DC.ToString, sDemoName + "demo", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDemoStart + DC.ToString, sDemoStart + "$", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDirArt + DC.ToString, sDirArt, N.Tag, Color.DarkGreen, Fnt)
		If DC = 1 Then
			AddNode(N, sZP + DC.ToString, sZP + "$02", N.Tag, Color.DarkGreen, Fnt)
			AddNode(N, sPacker + DC.ToString, sPacker + IIf(My.Settings.DefaultPacker = 1, "faster", "better"), N.Tag, Color.DarkGreen, Fnt)
		End If

		AddNewPartNode(N)       '[Add new part...]

        UpdateNewPartNode()     '->[Part 1] + [Add new file...]

        AddNewDiskNode()        '[Add new disk...]

        tv.Enabled = True
        tv.SelectedNode = N
        tv.EndUpdate()
        tv.ExpandAll()
        tv.Focus()

        CalcDiskNodeSize(N)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddNode(Parent As TreeNode, Name As String, Text As String, Optional Tag As Integer = 0, Optional NodeColor As Color = Nothing, Optional NodeFnt As Font = Nothing)
        On Error GoTo Err

        Parent.Nodes.Add(Name, Text)
        Parent.Nodes(Name).Tag = Tag

        Parent.Nodes(Name).ForeColor = If(NodeColor = Nothing, Color.Black, NodeColor)

        If NodeFnt IsNot Nothing Then
            Parent.Nodes(Name).NodeFont = NodeFnt
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub UpdateNode(Node As TreeNode, Text As String, Optional Tag As Integer = 0, Optional NodeColor As Color = Nothing, Optional NodeFnt As Font = Nothing)
        On Error GoTo Err

        With Node
            .Text = Text
            .Tag = Tag
            .ForeColor = NodeColor
            .NodeFont = NodeFnt
        End With

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddNewDiskNode()
        On Error GoTo Err

        tv.Nodes.Add(sAddDisk, "[Add new disk]")
        tv.Nodes(sAddDisk).Tag = 0                     'There is only ONE AddDisk node, its tag=0
        tv.Nodes(sAddDisk).ForeColor = Color.DarkBlue

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddNewPartNode(DiskNode As TreeNode)
        On Error GoTo Err

        Dim NPID As String = sAddPart + DC.ToString

        'Names of Add Part nodes are unique, tag is the same of the disk node's tag, to identify disk easily
        DiskNode.Nodes.Add(NPID, "[Add new part to this disk]")
        DiskNode.Nodes(NPID).Tag = DiskNode.Tag
        DiskNode.Nodes(NPID).ForeColor = Color.DarkBlue
        tv.SelectedNode = DiskNode.Nodes(NPID)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddNewFileNode(PartNode As TreeNode)
        On Error GoTo Err

        Dim NFID As String = sAddFile + CurrentPart.ToString

        'Names of Add File nodes are unique, tag is the same of the part node's tag, to identify disk and part easily
        PartNode.Nodes.Add(NFID, "[Add new file to this part]")
        PartNode.Nodes(NFID).Tag = PartNode.Tag
        PartNode.Nodes(NFID).ForeColor = Color.DarkBlue

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub UpdateNewPartNode()
        On Error GoTo Err

        tv.Enabled = False
        If Loading = False Then tv.BeginUpdate()
        PC += 1
        CurrentPart = PC
        With tv.SelectedNode
            .Text = "[Part " + PC.ToString + "]"
            .Name = .Parent.Name + ":P" + PC.ToString
            .Tag = .Parent.Tag + PC * &H1000
            .ForeColor = Color.DarkMagenta
            AddNewFileNode(tv.SelectedNode)
            AddNewPartNode(tv.SelectedNode.Parent)          'SelNode=[New part} node
            .Expand()
        End With
        If Loading = False Then tv.EndUpdate()
        tv.Enabled = True
        tv.Focus()
        tv.SelectedNode = tv.SelectedNode.Parent.Nodes(tv.SelectedNode.Parent.Name + ":P" + PC.ToString) 'Selnode=CurrentPart Node

        BtnPartUp.Enabled = True

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub UpdateFileNode()
        On Error GoTo Err

        Dim N As TreeNode = tv.SelectedNode

        FileType = 2  'Prg file

        OpenDemoFile()

        If NewFile = "" Then Exit Sub

        If Loading = False Then tv.BeginUpdate()

        FAddr = -1
        FOffs = -1
        FLen = 0

        If N.Index = 0 Then
            BlockCnt = 0
        Else
            BlockCnt = 1    'Fake block for compression
        End If

        N.Text = NewFile

        GetDefaultFileParameters(N)

        FAddr = DFAN
        FOffs = DFON
        FLen = DFLN
        DFA = True
        DFO = True
        DFL = True

        N.Text = CalcFileSize(N.Text, DFAN, DFLN)

        UpdateFileParameters(N)

        FileNameA(CurrentFile - 1) = N.Text
        FileSizeA(CurrentFile - 1) = FileSize

        CalcPartSize(N.Parent)

        If Loading = False Then tv.EndUpdate()

        tv.Focus()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddNewFile()
        On Error GoTo Err

        Dim N As TreeNode = tv.SelectedNode

        FileType = 2  'Prg file
        OpenDemoFile()

        If NewFile = "" Then Exit Sub

        FAddr = -1
        FOffs = -1
        FLen = 0

        If Loading = False Then tv.BeginUpdate()

        FC += 1
        ReDim Preserve FileNameA(FC - 1), FileAddrA(FC - 1), FileOffsA(FC - 1), FileLenA(FC - 1)

        If FileSizeA.Count < FC Then
            ReDim Preserve FileSizeA(FC - 1)
            ReDim Preserve FBSDisk(FC - 1)
        End If

        CurrentFile = FC

        If N.Index = 0 Then     'Is this the first file in this part?
            BlockCnt = 0        'Yes, reset block count
        Else
            BlockCnt = 1        'Fake block for compression
        End If

        With N
            .Text = NewFile
            .Name = .Parent.Name + ":F" + FC.ToString
            .Tag = .Parent.Tag + FC
            .ForeColor = Color.Black
        End With

        GetDefaultFileParameters(N)

        FAddr = DFAN
        FOffs = DFON
        FLen = DFLN

        DFA = True
        DFO = True
        DFL = True

        N.Text = CalcFileSize(N.Text, DFAN, DFLN)          'Also sets/clears IOBit

        UpdateFileParameters(N)

        FileSizeA(FC - 1) = FileSize
        FBSDisk(FC - 1) = CurrentDisk

        If PartSizeA.Count < CurrentPart Then
            ReDim Preserve PartSizeA(CurrentPart)
        End If

        AddNewFileNode(N.Parent)    'Add [Add New File] node before calculating part size

        CalcPartSize(N.Parent)

        If Loading = False Then tv.EndUpdate()
        tv.Focus()

        BtnFileUp.Enabled = True

        If CurrentDisk > 0 Then
            TssDisk.Text = "Disk " + (CurrentDisk).ToString + ": " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s free", " free")
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub UpdateDirArtPath()
        On Error GoTo Err

        Dim N As TreeNode = tv.SelectedNode

        FileType = 1  'Text file

        OpenDemoFile()

        If NewFile = "" Then Exit Sub

        N.Text = sDirArt + NewFile

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub UpdateDiskPath()
        On Error GoTo Err

        Dim N As TreeNode = tv.SelectedNode

        FileType = 3  'D64 file

        OpenDemoFile()

        If NewFile = "" Then Exit Sub

        N.Text = sDiskPath + NewFile

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub OpenDemoFile()
        On Error GoTo Err

        Dim P, F As String

        Dim N As TreeNode = tv.SelectedNode

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
                OpenFile("Open DirArt text File", "Text Files (*.txt)|*.txt", P)
            Case 2  'Prg File
                OpenFile("Open C64 Program File", "All Files (*.*)|*.*|PRG, SID, and Binary Files (*.prg; *.sid; *.bin)|*.prg; *.sid; *.bin|PRG Files (*.prg)|*.prg|Binary Files (*.bin)|*.bin|SID Files (*.sid)|*.sid", P)
            Case 3  'D64 file
                SaveFile("Save D64 Disk Image File", " D64 Files (*.d64)|*.d64", P, F, False)
        End Select

        Exit Sub
Err:
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

            If R = Windows.Forms.DialogResult.OK Then
                NewFile = .FileName
                FilePath = NewFile
            ElseIf R = Windows.Forms.DialogResult.Cancel Then
                NewFile = ""
            End If
        End With

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnNew_Click(sender As Object, e As EventArgs) Handles btnNew.Click
        On Error GoTo Err

        If MsgBox("Do you really want to start a new script?", vbQuestion + vbYesNo + vbDefaultButton2, "New script?") = vbNo Then Exit Sub

        SetScriptPath("")

        tssLabel.Text = "Script: (new script)"

        ResetArrays()

        DC = 0
        PC = 0
        FC = 0

        CurrentDisk = DC
        CurrentPart = PC
        CurrentFile = FC

        ReDim PartSizeA(PC), PartByteCntA(PC), PartBitCntA(PC), PartBitPosA(PC)

        PartByteCntA(PC) = 254
        PartBitCntA(PC) = 0
        PartBitPosA(PC) = 15

        DiskCnt = DC
        ReDim DiskSizeA(DiskCnt)

        If Loading = False Then tv.BeginUpdate()

        tv.Nodes.Clear()
        AddNewDiskNode()
        tv.SelectedNode = tv.Nodes(sAddDisk)
        BlankDiskStructure()

        If Loading = False Then tv.EndUpdate()
        tv.ExpandAll()
        tv.Focus()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        On Error GoTo Err

        OpenFile("Sparkle Loader Script files", "Sparkle Loader Script files (*.sls)|*.sls")
        If NewFile <> "" Then
            SetScriptPath(NewFile)

            OpenScript()
            Tv_GotFocus(sender, e)
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub OpenScript()
        On Error GoTo Err

        tssLabel.Text = "Script: " + ScriptName

        Script = IO.File.ReadAllText(ScriptName)

        ConvertScriptToNodes()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        On Error GoTo Err

        ConvertNodesToScript()

        If Script = "" Then
            Exit Sub
        End If

        SaveFile("Save Sparkle Loader Script files")

        If NewFile <> "" Then
            SetScriptPath(NewFile)

            SimplifyScript()

            tssLabel.Text = "Script: " + ScriptName
            IO.File.WriteAllText(ScriptName, Script)
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        On Error GoTo Err

        bBuildDisk = True

        Me.Close()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

	Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
		On Error GoTo Err

		bBuildDisk = False

		Me.Close()

		Exit Sub
Err:
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

			If R = Windows.Forms.DialogResult.OK Then
				NewFile = .FileName
			Else
				NewFile = ""
			End If
		End With

		Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub CalcPartSize(PartNode As TreeNode)
        On Error GoTo Err

        Cursor = Cursors.WaitCursor

        Dim FN As String
        Dim FA, FO, FL As String
        Dim FON, FLN As Integer
        Dim FUIO As Boolean = False

        Dim PNT(PartNode.Parent.Nodes.Count - 2) As String

        Dim NI As Integer
        For I As Integer = 0 To PartNode.Parent.Nodes.Count - 1
            If InStr(PartNode.Parent.Nodes(I).Text, "[Part") <> 0 Then
                NI = I
                Exit For
            End If
        Next

        For I As Integer = NI To PartNode.Parent.Nodes.Count - 2
            PNT(I) = PartNode.Parent.Nodes(I).Text
        Next

        Dim P() As Byte

        If PartNode.Nodes.Count = 1 Then
            CurrentPart = PartNode.Index - (NI - 1)
            PartNode.Text = "[Part " + CurrentPart.ToString + "]"
            PartNode.Tag = PartNode.Parent.Tag + CurrentPart * &H1000
            PartSizeA(CurrentPart - 1) = 0
            tv.Refresh()
            GoTo Done
        End If

        For K = PartNode.Index To PartNode.Parent.Nodes.Count - 2

            PartNode = PartNode.Parent.Nodes(K)
            If PartNode.Nodes.Count > 1 Then

                FileCnt = -1
                ReDim FileNameA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt), FileIOA(FileCnt)
                Prgs.Clear()
                ReDim ByteSt(-1)
                UncomPartSize = 0
                BlockCnt = 0

                For I As Integer = 0 To PartNode.Nodes.Count - 2
                    FN = PartNode.Nodes(I).Text

                    If Strings.Right(FN, 1) = "*" Then
                        FN = Replace(FN, "*", "")
                        FUIO = True
                    Else
                        FUIO = False
                    End If

                    If IO.File.Exists(FN) = True Then
                        P = IO.File.ReadAllBytes(FN)

                        FA = Strings.Right(PartNode.Nodes(I).Nodes(0).Text, 4)
                        FO = Strings.Right(PartNode.Nodes(I).Nodes(1).Text, 4)
                        FL = Strings.Right(PartNode.Nodes(I).Nodes(2).Text, 4)

                        FON = Convert.ToInt32(FO, 16)
                        FLN = Convert.ToInt32(FL, 16)

                        'Make sure file length is not longer than actual file (should not happen)
                        If FON + FLN > P.Length Then
                            FLN = P.Length - FON
                            FL = ConvertNumberToHexString(FLN Mod 256, Int(FLN / 256))
                        End If

                        UncomPartSize += Int(FLN / 254)
                        If FLN Mod 254 <> 0 Then
                            UncomPartSize += 1
                        End If

                        'Trim file to the specified data segment (FLN number of bytes starting at FON, to Address of FAN)
                        For J As Integer = 0 To FLN - 1
                            P(J) = P(FON + J)
                        Next
                        ReDim Preserve P(FLN - 1)

                    Else

                        MsgBox("The following file does not exist:" + vbNewLine + vbNewLine + FN)
                        GoTo NoDisk
                    End If

                    FileCnt += 1
                    ReDim Preserve FileNameA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt), FileIOA(FileCnt)

                    FileNameA(FileCnt) = FN
                    FileAddrA(FileCnt) = FA
                    FileOffsA(FileCnt) = FO     'This may not be needed later
                    FileLenA(FileCnt) = FL
                    FileIOA(FileCnt) = FUIO

                    Prgs.Add(P)
                Next

                CurrentPart = Int((PartNode.Tag And &HFFF000) / &H1000)
                PartCnt = CurrentPart

                BufferCnt = 0
                ByteCnt = PartByteCntA(CurrentPart - 1)
                BitCnt = PartBitCntA(CurrentPart - 1)
                BitPos = PartBitPosA(CurrentPart - 1)

                SortPart()
                CompressPart(True)

                If CurrentPart + 1 > PartByteCntA.Count Then
                    ReDim Preserve PartByteCntA(CurrentPart + 1), PartBitCntA(CurrentPart + 1), PartBitPosA(CurrentPart + 1)
                End If

                PartByteCntA(CurrentPart) = ByteCnt
                PartBitCntA(CurrentPart) = BitCnt
                PartBitPosA(CurrentPart) = BitPos

                For I As Integer = 0 To PartNode.Parent.Nodes.Count - 1
                    'Find the first part node under this disk node
                    If Strings.Left(PartNode.Parent.Nodes(I).Text, 5) = "[Part" Then
                        'If our current part node is the first one under this disk, then increase BufferCnt
                        If PartNode.Index = I Then
                            BufferCnt += 1
                        End If
                        Exit For
                    End If
                Next

                PartSizeA(CurrentPart - 1) = BufferCnt
                UncomPartSize = Int(10000 * BufferCnt / UncomPartSize) / 100

                If Prgs.Count = 0 Then
                    PNT(K) = "[Part " + (K - 5).ToString + "]"
                Else
                    PNT(K) = "[Part " + (K - 5).ToString + ": " + BufferCnt.ToString + " block" + IIf(BufferCnt <> 1, "s", "") + " compressed, " _
                + UncomPartSize.ToString + "% of uncompressed size]"
                End If
            End If
        Next

        If Loading = False Then tv.BeginUpdate()
        For I As Integer = NI To PartNode.Parent.Nodes.Count - 2
            If PartNode.Parent.Nodes(I).Text <> PNT(I) Then
                PartNode.Parent.Nodes(I).Text = PNT(I)
            End If
        Next

        If Loading = False Then tv.EndUpdate()

Done:
        CalcDiskNodeSize(PartNode.Parent)

        If Loading = False Then Cursor = Cursors.Default

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        If Loading = False Then Cursor = Cursors.Default

    End Sub

    Private Function CalcFileSize(FN As String, FA As Integer, FL As Integer) As String
        On Error GoTo Err

        CalcFileSize = FN

        Dim DefaultFUIO As Boolean = InStr(FN, "*") <> 0

        FAddr = FA
        FLen = FL

        FileSize = CalcOrigBlockCnt()   'This also opens the prg to Prg() and calculates FAddr and FLen

        FN = Replace(FN, "*", "")

        If UnderIO() = True Then
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
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Sub SwapIOStatus()

        With SelNode
            If Strings.Right(.Text, 3) = "yes" Then
                .Text = sFileUIO + " no"
                .ForeColor = Color.MediumPurple
                .Parent.Text = Strings.Replace(.Parent.Text, "*", "")
                CalcPartSize(.Parent.Parent)
            ElseIf Strings.Right(.Text, 3) = " no" Then
                FAddr = Convert.ToInt32(Strings.Right(.Parent.Nodes(0).Text, 4), 16)
                FLen = Convert.ToInt32(Strings.Right(.Parent.Nodes(2).Text, 4), 16)
                If UnderIO() Then
                    .Text = sFileUIO + "yes"
                    .ForeColor = Color.Purple
                    .Parent.Text = Strings.Replace(.Parent.Text, "*", "") + "*"
                    CalcPartSize(.Parent.Parent)
                End If
            Else
            End If
        End With

    End Sub

    Private Function CalcOrigBlockCnt() As Integer
        On Error GoTo Err

        Dim Ext As String

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
                    If FLen <> Prg.Length - FOffs Then
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
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Sub BtnFileDown_Click(sender As Object, e As EventArgs) Handles BtnFileDown.Click
        On Error GoTo Err

        If tv.SelectedNode Is Nothing Then Exit Sub

        Dim N As TreeNode = tv.SelectedNode
        Dim P As TreeNode = N.Parent
        Dim I As Integer = N.Index

        If I < P.Nodes.Count - 2 Then
            If Loading = False Then tv.BeginUpdate()
            P.Nodes.RemoveAt(I)
            P.Nodes.Insert(I + 1, N)
            If Loading = False Then tv.EndUpdate()
            tv.SelectedNode = N
            tv.Focus()
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnFileUp_Click(sender As Object, e As EventArgs) Handles BtnFileUp.Click
        On Error GoTo Err

        If tv.SelectedNode Is Nothing Then Exit Sub

        Dim N As TreeNode = tv.SelectedNode
        Dim P As TreeNode = N.Parent
        Dim I As Integer = N.Index

        If I > 0 Then
            If Loading = False Then tv.BeginUpdate()
            P.Nodes.RemoveAt(I)
            P.Nodes.Insert(I - 1, N)
            If Loading = False Then tv.EndUpdate()
            tv.SelectedNode = N
            tv.Focus()
        End If

        CalcPartSize(P)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnPartDown_Click(sender As Object, e As EventArgs) Handles BtnPartDown.Click
        On Error GoTo Err

        If tv.SelectedNode Is Nothing Then Exit Sub

        Dim N As TreeNode = tv.SelectedNode
        Dim P As TreeNode = N.Parent
        Dim I As Integer = N.Index
        Dim Name1 As String = N.Name
        Dim Tag1 As Integer = N.Tag

        If I < P.Nodes.Count - 2 Then
            Dim Name2 As String = P.Nodes(I + 1).Name
            Dim Tag2 As Integer = P.Nodes(I + 1).Tag
            N.Name = Name2
            N.Tag = Tag2
            P.Nodes(I + 1).Name = Name1
            P.Nodes(I + 1).Tag = Tag1
            If Loading = False Then tv.BeginUpdate()
            P.Nodes.RemoveAt(I)
            P.Nodes.Insert(I + 1, N)
            If Loading = False Then tv.EndUpdate()
            tv.SelectedNode = N
            tv.Focus()
        End If

        CalcPartSize(P.Nodes(I))

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub BtnPartUp_Click(sender As Object, e As EventArgs) Handles BtnPartUp.Click
        On Error GoTo Err

        If tv.SelectedNode Is Nothing Then Exit Sub

        Dim N As TreeNode = tv.SelectedNode
        Dim P As TreeNode = N.Parent
        Dim I As Integer = N.Index
        Dim Name1 As String = N.Name
        Dim Tag1 As Integer = N.Tag

        If I > 0 Then
            Dim Name2 As String = P.Nodes(I - 1).Name
            Dim Tag2 As Integer = P.Nodes(I - 1).Tag
            N.Name = Name2
            N.Tag = Tag2
            P.Nodes(I - 1).Name = Name1
            P.Nodes(I - 1).Tag = Tag1
            If Loading = False Then tv.BeginUpdate()
            P.Nodes.RemoveAt(I)
            P.Nodes.Insert(I - 1, N)
            If Loading = False Then tv.EndUpdate()
            tv.SelectedNode = N
            tv.Focus()
        End If

        CalcPartSize(N)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ToggleFileNodes()
        On Error GoTo Err

        If Loading = False Then tv.BeginUpdate()

        If chkExpand.Checked = False Then
            If tv.Nodes.Count > 1 Then
                For D As Integer = 0 To tv.Nodes.Count - 2
                    If tv.Nodes(D).Nodes.Count > 1 Then
                        For P As Integer = 0 To tv.Nodes(D).Nodes.Count - 2
                            If tv.Nodes(D).Nodes(P).Nodes.Count > 1 Then
                                For F As Integer = 0 To tv.Nodes(D).Nodes(P).Nodes.Count - 2
                                    tv.Nodes(D).Nodes(P).Nodes(F).Collapse()
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Else
            If tv.Nodes.Count > 1 Then
                For D As Integer = 0 To tv.Nodes.Count - 2
                    If tv.Nodes(D).Nodes.Count > 1 Then
                        For P As Integer = 0 To tv.Nodes(D).Nodes.Count - 2
                            If tv.Nodes(D).Nodes(P).Nodes.Count > 1 Then
                                For F As Integer = 0 To tv.Nodes(D).Nodes(P).Nodes.Count - 2
                                    tv.Nodes(D).Nodes(P).Nodes(F).Expand()
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        End If

        If Loading = False Then tv.EndUpdate()
        tv.Focus()
        If tv.SelectedNode IsNot Nothing Then
            tv.SelectedNode.EnsureVisible()
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function UnderIO() As Boolean
        On Error GoTo Err

        If (FAddr >= &HD000) And (FAddr < &HE000) Then
            'File beings under IO
            UnderIO = True
        ElseIf (FAddr + FLen - 1 >= &HD000) And (FAddr + FLen - 1 < &HE000) Then
            'File ends under IO
            UnderIO = True
        ElseIf (FAddr < &HD000) And (FAddr + FLen - 1 >= &HE000) Then
            'File begins before and ends after IO
            UnderIO = True
        Else
            'Otherwise
            UnderIO = False
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Sub UpdateFileParameters(FileNode As TreeNode)
        On Error GoTo Err

        If FileNode Is Nothing Then Exit Sub

        If Loading = False Then tv.BeginUpdate()

        If FileNode.Nodes(FileNode.Name + ":FA") Is Nothing Then
            FileNode.Nodes.Add(FileNode.Name + ":FA", sFileAddr + ConvertNumberToHexString(FAddr Mod 256, Int(FAddr / 256)))
            With FileNode.Nodes(FileNode.Name + ":FA")
                .Tag = FileNode.Tag
                .ForeColor = IIf(DFA = True, DefaultCol, ManualCol)
                .NodeFont = New Font("Consolas", 10)
            End With
        Else
            FileNode.Nodes(FileNode.Name + ":FA").Text = sFileAddr + ConvertNumberToHexString(FAddr Mod 256, Int(FAddr / 256))
        End If

        If FileNode.Nodes(FileNode.Name + ":FO") Is Nothing Then
            FileNode.Nodes.Add(FileNode.Name + ":FO", sFileOffs + ConvertNumberToHexString(FOffs Mod 256, Int(FOffs / 256)))
            With FileNode.Nodes(FileNode.Name + ":FO")
                .Tag = FileNode.Tag
                .ForeColor = IIf(DFO = True, DefaultCol, ManualCol)
                .NodeFont = New Font("Consolas", 10)
            End With
        Else
            FileNode.Nodes(FileNode.Name + ":FO").Text = sFileOffs + ConvertNumberToHexString(FOffs Mod 256, Int(FOffs / 256))
        End If

        If FileNode.Nodes(FileNode.Name + ":FL") Is Nothing Then
            FileNode.Nodes.Add(FileNode.Name + ":FL", sFileLen + ConvertNumberToHexString(FLen Mod 256, Int(FLen / 256)))
            With FileNode.Nodes(FileNode.Name + ":FL")
                .Tag = FileNode.Tag
                .ForeColor = IIf(DFL = True, DefaultCol, ManualCol)
                .NodeFont = New Font("Consolas", 10)
            End With
        Else
            FileNode.Nodes(FileNode.Name + ":FL").Text = sFileLen + ConvertNumberToHexString(FLen Mod 256, Int(FLen / 256))
        End If

        If FileNode.Nodes(FileNode.Name + ":FUIO") Is Nothing Then
            FileNode.Nodes.Add(FileNode.Name + ":FUIO", sFileUIO + IIf(FileUnderIO = True, "yes", IIf(UnderIO() = True, " no", "n/a")))
            With FileNode.Nodes(FileNode.Name + ":FUIO")
                .Tag = FileNode.Tag
                .ForeColor = IIf(FileUnderIO = True, Color.Purple, Color.MediumPurple)
                .NodeFont = New Font("Consolas", 10)
            End With
        Else
            FileNode.Nodes(FileNode.Name + ":FUIO").Text = sFileUIO + IIf(FileUnderIO = True, "yes", IIf(UnderIO() = True, " no", "n/a"))
        End If

        If FileNode.Nodes(FileNode.Name + ":FS") Is Nothing Then
            FileNode.Nodes.Add(FileNode.Name + ":FS", sFileSize + FileSize.ToString + " block" + IIf(FileSize <> 1, "s", ""))
            With FileNode.Nodes(FileNode.Name + ":FS")
                .Tag = FileNode.Tag
                .ForeColor = Color.DarkGray
                .NodeFont = New Font("Consolas", 10)
            End With
        Else
            FileNode.Nodes(FileNode.Name + ":FS").Text = sFileSize + FileSize.ToString + " block" + IIf(FileSize <> 1, "s", "")
        End If

        If Loading = False Then tv.EndUpdate()

        FileNode.Expand()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function ConvertScriptToNodes() As Boolean
        On Error GoTo Err

        ConvertScriptToNodes = True

        Cursor = Cursors.WaitCursor

        SS = 1 : SE = 1

        FindNextScriptEntry()

        If ScriptEntry <> ScriptHeader Then
            MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
            Exit Function
        End If

        Loading = True

        tv.Enabled = False
        'tv.ShowNodeToolTips = False
        tv.BeginUpdate()

        DC = 0
        PC = 0
        FC = 0

        CurrentDisk = DC
        CurrentPart = PC
        CurrentFile = FC

        ReDim PartSizeA(PC), PartByteCntA(PC), PartBitCntA(PC), PartBitPosA(PC)

        PartByteCntA(PC) = 254
        PartBitCntA(PC) = 0
        PartBitPosA(PC) = 15

        DiskCnt = DC
        ReDim DiskSizeA(DiskCnt)

        tv.Nodes.Clear()
        AddNewDiskNode()

		Packer = My.Settings.DefaultPacker

NewDisk:
        tv.SelectedNode = tv.Nodes(sAddDisk)

        'Reset buffer and other disk variables here
        ResetDiskVariables()

        BlankDiskStructure()

        Dim DiskNode As TreeNode = tv.SelectedNode

FindNext:
        FindNextScriptEntry()

        If InStr(ScriptEntry, vbTab) = 0 Then
            ScriptEntryType = ScriptEntry
        Else
            ScriptEntryType = Strings.Left(ScriptEntry, InStr(ScriptEntry, vbTab) - 1)
            ScriptEntry = Strings.Right(ScriptEntry, ScriptEntry.Length - InStr(ScriptEntry, vbTab))
        End If

        SplitEntry()

        Select Case LCase(ScriptEntryType)
            Case "path:"
                If InStr(ScriptEntryArray(0), ":") = 0 Then
                    ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
                End If
                Dim Fnt As New Font("Consolas", 10)
                UpdateNode(DiskNode.Nodes(sDiskPath + DC.ToString), sDiskPath + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt)', tDiskPath)
            Case "header:"
                Dim Fnt As New Font("Consolas", 10)
                UpdateNode(DiskNode.Nodes(sDiskHeader + DC.ToString), sDiskHeader + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt)', tDiskHeader)
            Case "id:"
                Dim Fnt As New Font("Consolas", 10)
                UpdateNode(DiskNode.Nodes(sDiskID + DC.ToString), sDiskID + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt)', tDiskID)
            Case "name:"
                Dim Fnt As New Font("Consolas", 10)
                UpdateNode(DiskNode.Nodes(sDemoName + DC.ToString), sDemoName + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt)', tDemoName)
            Case "start:"
                Dim Fnt As New Font("Consolas", 10)
                UpdateNode(DiskNode.Nodes(sDemoStart + DC.ToString), sDemoStart + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, Color.DarkGreen, Fnt)', tDemoStart)
            Case "dirart:"
                Dim Fnt As New Font("Consolas", 10)
                If ScriptEntryArray(0) <> "" Then
                    If InStr(ScriptEntryArray(0), ":") = 0 Then
                        ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
                    End If
                End If
				UpdateNode(DiskNode.Nodes(sDirArt + DC.ToString), sDirArt + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt)
			Case "zp:"
				If CurrentDisk = 1 Then 'ZP can only be set from the first disk
					Dim Fnt As New Font("Consolas", 10)
					UpdateNode(DiskNode.Nodes(sZP + DC.ToString), sZP + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, Color.DarkGreen, Fnt)
				End If
			Case "packer:"
				If CurrentDisk = 1 Then 'ZP can only be set from the first disk
					Dim Fnt As New Font("Consolas", 10)
					UpdateNode(DiskNode.Nodes(sPacker + DC.ToString), sPacker + LCase(ScriptEntryArray(0)), DiskNode.Tag, Color.DarkGreen, Fnt)
					Select Case LCase(ScriptEntryArray(0))
						Case "faster"
							Packer = 1
						Case "better"
							Packer = 2
					End Select
				End If
			Case "file:"
				AddFileFromScript(DiskNode)
            Case "new disk"

                'Update last part size and disk size before starting new disk node
                'If DiskNode.Nodes.Count > 7 Then UpdatePartSize(DiskNode)   'First 6 nodes are disk info, last one is AddPart node
                UpdatePartSize(DiskNode)
                CurrentDisk = DiskCnt
                DiskNode.Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk - 1).ToString + " block" + IIf(DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " used, " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " free]"

                GoTo NewDisk
            Case Else
                'Figure out what to do with comments here...
        End Select

        If SE < Script.Length Then GoTo FindNext

        'Update last part size and disk size before finishing
        UpdatePartSize(DiskNode)
        CurrentDisk = DiskCnt
        DiskNode.Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk - 1).ToString + " block" + IIf(DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " used, " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " free]"

        GoTo Done
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        ConvertScriptToNodes = False
Done:
        With tv
            .EndUpdate()
            .Enabled = True
            ToggleFileNodes()
            .Focus()
            .SelectedNode = tv.Nodes(0)
            .SelectedNode.EnsureVisible()
        End With

        Loading = False

        Cursor = Cursors.Default

    End Function

    Private Function AddFileFromScript(DiskNode As TreeNode) As Boolean
        On Error GoTo Err

        AddFileFromScript = True

        FC += 1
        CurrentFile = FC

        tv.SelectedNode = DiskNode.Nodes(DiskNode.Name + ":P" + PC.ToString)

        If (NewPart = True) And (Prgs.Count > 0) Then
            UpdatePartSize(DiskNode)

            NewPart = False
            tv.SelectedNode = DiskNode.Nodes(sAddPart + DC.ToString)
            PC += 1
            CurrentPart = PC
            With tv.SelectedNode
                .Name = .Parent.Name + ":P" + PC.ToString
                .Tag = .Parent.Tag + PC * &H1000
                .ForeColor = Color.DarkMagenta
                AddNewFileNode(tv.SelectedNode)
                AddNewPartNode(tv.SelectedNode.Parent)          'SelNode=[New part} node
                .Expand()
            End With
            tv.SelectedNode = tv.SelectedNode.Parent.Nodes(tv.SelectedNode.Parent.Name + ":P" + PC.ToString) 'SelNode=CurrentPart Node
        Else
            tv.SelectedNode = DiskNode.Nodes(DiskNode.Name + ":P" + PC.ToString)
        End If

        AddFileToPart()     'This will check and correct file parameters to "0000" format

        Dim N As TreeNode = tv.SelectedNode.Nodes(sAddFile + PC.ToString)

        Dim FilePath As String = ScriptEntryArray(0)

        If InStr(FilePath, ":") = 0 Then
            FilePath = ScriptPath + FilePath
        End If

        With N
            .Text = FilePath                                 'New File's Path
            .Name = N.Parent.Name + ":F" + FC.ToString
            .ForeColor = Color.Black
            .Tag = .Parent.Tag + FC
        End With

        Dim Ext As String
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

        FileSize = Int(FLen / 256)
        If FLen Mod 256 <> 0 Then FileSize += 1

        UpdateFileParameters(N)

        AddNewFileNode(N.Parent)

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        AddFileFromScript = False

    End Function

    Private Sub UpdatePartSize(DiskNode As TreeNode)
        On Error GoTo Err

        SortPart()

        BufferCnt = 0
        ByteCnt = PartByteCntA(PC - 1)
        BitCnt = PartBitCntA(PC - 1)
        BitPos = PartBitPosA(PC - 1)

        If CompressPart(True) = False Then Exit Sub

        'Save current parts compressed size to array
        ReDim Preserve PartSizeA(PC)
        PartSizeA(PC - 1) = BufferCnt

        For I As Integer = 0 To DiskNode.Nodes.Count - 1
            'Find the first part node under this disk node
            If Strings.Left(DiskNode.Nodes(I).Text, 5) = "[Part" Then
                'If our current part node is the first one on this disk, then increase BufferCnt
                If DiskNode.Nodes(DiskNode.Name + ":P" + PC.ToString).Index = I Then
                    If Prgs.Count <> 0 Then BufferCnt += 1
                End If
                Exit For
            End If
        Next

        UncomPartSize = Int(10000 * BufferCnt / UncomPartSize) / 100

        tv.SelectedNode = DiskNode.Nodes(DiskNode.Name + ":P" + PC.ToString)

        If Loading = False Then tv.BeginUpdate()

        If Prgs.Count = 0 Then
            tv.SelectedNode.Text = "[Part " + PC.ToString + "]"
        Else
            tv.SelectedNode.Text = "[Part " + PC.ToString + ": " + BufferCnt.ToString + " block" + IIf(BufferCnt <> 1, "s", "") + " compressed, " _
               + UncomPartSize.ToString + "% of uncompressed size]"
        End If

        If Loading = False Then tv.EndUpdate()

        DiskSizeA(DiskCnt - 1) += BufferCnt

        'Save next part's buffer variables to arrays
        ReDim Preserve PartByteCntA(PC + 1), PartBitCntA(PC + 1), PartBitPosA(PC + 1)
        PartByteCntA(PC) = ByteCnt
        PartBitCntA(PC) = BitCnt
        PartBitPosA(PC) = BitPos

        ResetPartVariables()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ConvertNodesToScript()
        On Error GoTo Err

        Dim S As String = ScriptHeader + vbNewLine + vbNewLine
        Dim SA, DI, DH As String
        Dim DP As Integer = 7   'Default first part node index (7 for the first disk, 6 for the rest)

        For D As Integer = 0 To tv.Nodes.Count - 2
            Dim N As TreeNode = tv.Nodes(D)

            SA = "" 'Determine demo start address
            If N.Nodes(4).Text.Length > (sDemoStart.Length + 1) Then
                SA = Strings.Right(N.Nodes(4).Text, N.Nodes(4).Text.Length - Len(sDemoStart + "$"))
            Else
                If N.Nodes.Count > 6 Then
                    If N.Nodes(6).Nodes.Count > 1 Then
                        SA = Strings.Right(N.Nodes(6).Nodes(0).Nodes(0).Text, 4)
                    End If
                End If
            End If

            'Determine disk header
            DH = Strings.Right(N.Nodes(1).Text, N.Nodes(1).Text.Length - Len(sDiskHeader))
            If DH = "" Then
                DH = "demo disk " + Year(Now).ToString
            End If

            'Determine disk ID
            DI = Strings.Right(N.Nodes(2).Text, N.Nodes(2).Text.Length - Len(sDiskID))
            If DI = "" Then
                DI = "sprkl"
            End If

            S += "Path:" + vbTab + Strings.Right(N.Nodes(0).Text, N.Nodes(0).Text.Length - Len(sDiskPath)) + vbNewLine +
            "Header:" + vbTab + DH + vbNewLine +
            "ID:" + vbTab + DI + vbNewLine +
            "Name:" + vbTab + Strings.Right(N.Nodes(3).Text, N.Nodes(3).Text.Length - Len(sDemoName)) + vbNewLine +
            "Start:" + vbTab + SA + vbNewLine +
            "DirArt:" + vbTab + Strings.Right(N.Nodes(5).Text, N.Nodes(5).Text.Length - Len(sDirArt))
            If D = 0 Then
				'ZP only for first disk
				S += vbNewLine + "ZP:" + vbTab + Strings.Right(N.Nodes(6).Text, N.Nodes(6).Text.Length - Len(sZP + "$"))
				S += vbNewLine + "Packer:" + vbTab + Strings.Right(N.Nodes(7).Text, N.Nodes(7).Text.Length - Len(sPacker))
			End If

            For P As Integer = DP To N.Nodes.Count - 2
                If N.Nodes(P).Nodes.Count > 1 Then
                    S += vbNewLine
                    For F As Integer = 0 To N.Nodes(P).Nodes.Count - 2
                        Dim FN As TreeNode = N.Nodes(P).Nodes(F)

                        S += vbNewLine + "File:" + vbTab + FN.Text

                        If FN.Nodes(0).ForeColor = ManualCol Then
                            S += vbTab + Strings.Right(FN.Nodes(0).Text, 4)
                        End If

                        If FN.Nodes(1).ForeColor = ManualCol Then
                            S += vbTab + Strings.Right(FN.Nodes(1).Text, 4)
                        End If

                        If FN.Nodes(2).ForeColor = ManualCol Then
                            S += vbTab + Strings.Right(FN.Nodes(2).Text, 4)
                        End If
                    Next
                End If
            Next
            If D < tv.Nodes.Count - 2 Then
                S += vbNewLine + vbNewLine + "New Disk" + vbNewLine + vbNewLine
            End If

            DP = 6

        Next

        Script = S

        'MsgBox(S)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub CalcDiskNodeSize(DiskNode As TreeNode)
        On Error GoTo Err

        If DiskNode Is Nothing Then Exit Sub

        CurrentDisk = Int(DiskNode.Tag / &H1000000)

        DiskSizeA(CurrentDisk - 1) = 0

        For I As Integer = 0 To DiskNode.Nodes.Count - 2
            CurrentPart = Int((DiskNode.Nodes(I).Tag And &HFFF000) / &H1000)
            If CurrentPart <> 0 Then
                If PartSizeA.Count > CurrentPart - 1 Then
                    DiskSizeA(CurrentDisk - 1) += PartSizeA(CurrentPart - 1)
                End If
            End If
        Next

        tv.Refresh()

        DiskNode.Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk - 1).ToString + " block" + IIf(DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " used, " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " free]"

        If DiskSizeA(CurrentDisk - 1) > 664 Then
            MsgBox("The size of this disk exceeds 664 blocks!", vbOKOnly + vbCritical, "Disk is full!")
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FrmSE_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        On Error GoTo Err

        If txtEdit.Visible Then tv.Focus()

		With My.Settings
			.DefaultPacker = IIf(OptFaster.Checked, 1, 2)
			.ShowFileDetails = chkExpand.Checked
			.ShowToolTips = chkToolTips.Checked
			.EditorWindowMax = Me.WindowState = FormWindowState.Maximized
			If WindowState = FormWindowState.Normal Then
				.EditorHeight = Height
				.EditorWidth = Width
			End If
		End With

		ConvertNodesToScript()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FrmSE_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        On Error GoTo Err

        Dim DropFiles() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each Path In DropFiles
            Select Case Strings.Right(Path, 4)
                Case ".sls"
                    SetScriptPath(Path)

                    OpenScript()
                    Tv_GotFocus(sender, e)
            End Select
        Next

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FrmSE_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        On Error GoTo Err

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

	Private Sub OptFaster_CheckedChanged(sender As Object, e As EventArgs) Handles OptFaster.CheckedChanged

		If OptFaster.Checked = True Then
			My.Settings.DefaultPacker = 1
		Else
			My.Settings.DefaultPacker = 2
		End If

	End Sub

	Private Sub Tv_DragDrop(sender As Object, e As DragEventArgs) Handles tv.DragDrop
		On Error GoTo Err

		FrmSE_DragDrop(sender, e)

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Tv_DragEnter(sender As Object, e As DragEventArgs) Handles tv.DragEnter
        On Error GoTo Err

        FrmSE_DragEnter(sender, e)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_BeforeExpand(sender As Object, e As TreeViewCancelEventArgs) Handles tv.BeforeExpand
        On Error GoTo Err

        If (tv.SelectedNode.Tag And &HFFF) <> 0 Then
            If chkExpand.Checked = False Then
                If Dbl = True Then e.Cancel = True
            End If
        End If

        Dbl = False

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_BeforeCollapse(sender As Object, e As TreeViewCancelEventArgs) Handles tv.BeforeCollapse
        On Error GoTo Err

        If (tv.SelectedNode.Tag And &HFFF) <> 0 Then
            If chkExpand.Checked = True Then
                If Dbl = True Then
                    e.Cancel = True
                End If
            End If
        End If

        Dbl = False

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ChkToolTips_CheckedChanged(sender As Object, e As EventArgs) Handles chkToolTips.CheckedChanged
        On Error GoTo Err

        If chkToolTips.Checked = False Then
            TT.Hide(txtEdit)
            TT.Hide(tv)
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ChkExpand_CheckedChanged(sender As Object, e As EventArgs) Handles chkExpand.CheckedChanged
        On Error GoTo Err

        ToggleFileNodes()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_MouseDown(sender As Object, e As MouseEventArgs) Handles tv.MouseDown
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
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FrmSE_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        On Error GoTo Err

        With tv
            .Width = Width - btnNew.Width - 56
            .Height = Height - .Top - strip.Height - 56
        End With

        With btnNew
            .Left = Width - .Width - 32
            btnLoad.Left = .Left
            btnSave.Left = .Left
            BtnPartUp.Left = .Left
            BtnPartDown.Left = .Left
            chkExpand.Left = .Left
            chkToolTips.Left = .Left
			PnlPacker.Left = .Left - 11
		End With

		With btnOK
            .Top = tv.Top + tv.Height - .Height
            .Left = btnNew.Left
        End With

        With btnCancel
            .Top = btnOK.Top - .Height - 6
            .Left = btnNew.Left
        End With


		Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FrmSE_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        On Error GoTo Err

        Dim R As Rectangle = Screen.FromPoint(Me.Location).WorkingArea

        Dim X = R.Left + (R.Width - Width) \ 2
        Dim Y = R.Top + (R.Height - Height) \ 2
        Location = New Point(X, Y)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub TxtEdit_MouseHover(sender As Object, e As EventArgs) Handles txtEdit.MouseHover
        On Error GoTo Err

        ShowTT()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ShowTT()
        On Error GoTo Err

        TT.Hide(txtEdit)

        If chkToolTips.Checked = False Then Exit Sub

        If Loading = True Then Exit Sub

        Dim TTT As String = ""

        With TT
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
                    TTT = "Type in the demo's entry point." +
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
                Case Else
                    Exit Select
            End Select

            If TTT <> "" Then
                'TT.Show(TTT, txtEdit, 5000)
                TT.Show(TTT, txtEdit)
            End If
        End With

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_NodeMouseHover(sender As Object, e As TreeNodeMouseHoverEventArgs) Handles tv.NodeMouseHover
        On Error GoTo Err

        TT.Hide(tv)

        If chkToolTips.Checked = False Then Exit Sub

        If Loading = True Then Exit Sub

        Dim TTT As String = ""

        With TT
            If Strings.Left(e.Node.Text, 13) = "[Add new disk" Then
                .ToolTipTitle = "Add New Demo Disk"
                TTT = tAddDisk
            ElseIf Strings.Left(e.Node.Text, 13) = "[Add new part" Then
                .ToolTipTitle = "Add New Demo Part"
                TTT = tAddPart
            ElseIf Strings.Left(e.Node.Text, 13) = "[Add new file" Then
                .ToolTipTitle = "Add New Demo File"
                TTT = tAddFile
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
            ElseIf Strings.Left(e.Node.Text, 5) = "[Part" Then
                .ToolTipTitle = "Demo Part"
                TTT = tPart
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
					Case sPacker
						.ToolTipTitle = "Packer to be used"
						TTT = tPacker
					Case Else
				End Select
            End If

            If TTT <> "" Then
                .Show(TTT, tv, 5000)
            End If
        End With

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SCC_CallBackProc(ByRef m As Message) Handles SCC.CallBackProc
        On Error GoTo Err

        'If txtEdit is visible while we are scrolling - set focus back to Tv and hide txtEdit
        Select Case m.Msg
            Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
                tv.Focus()
        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub Tv_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles tv.NodeMouseClick
        On Error GoTo Err

        If e.Button = MouseButtons.Right Then
            tv.SelectedNode = e.Node
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub TxtEdit_MouseWheel(sender As Object, e As MouseEventArgs) Handles txtEdit.MouseWheel
        On Error GoTo Err

        tv.Focus()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub
End Class