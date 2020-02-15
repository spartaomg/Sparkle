Public Class FrmEditor

    Private ReadOnly sDiskPath As String = "Disk Path: "
    Private ReadOnly sDiskHeader As String = "Disk Header: "
    Private ReadOnly sDiskID As String = "Disk ID: "
    Private ReadOnly sDemoName As String = "Demo Name: "
    Private ReadOnly sDemoStart As String = "Demo Start: "
    Private ReadOnly sAddEntry As String = "AddEntry"
    Private ReadOnly sAddScript As String = "AddScript"
    Private ReadOnly sAddDisk As String = "AddDisk"
    Private ReadOnly sAddPart As String = "AddPart"
    Private ReadOnly sAddFile As String = "AddFile"
    Private ReadOnly sFileSize As String = "Original File Size: "
    Private ReadOnly sFileUIO As String = "Load to:      "     '"Load under I/O: "
    Private ReadOnly sFileAddr As String = "Load Address: $"
    Private ReadOnly sFileOffs As String = "File Offset:  $"
    Private ReadOnly sFileLen As String = "File Length:  $"
    Private ReadOnly sDirArt As String = "DirArt: "
    Private ReadOnly sZP As String = "Zeropage: "
    Private ReadOnly sPacker As String = "Packer: "

    Private SC, DC, PC, FC As Integer

    Private ScriptNode, SelNode, NewScriptNode As TreeNode
    Private NewD As Boolean = False

    Private Sub FrmEditor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error GoTo Err

        ScriptNode = TV.Nodes.Add("ThisScript", "This script: " + If(ScriptPath <> "", ScriptPath + "Sparkle v1.3 2020", "(New Script)"))
        SelNode = ScriptNode
        ScriptNode.BackColor = Color.LightGray
        AddNewEntryNode(ScriptNode)

        TV.ExpandAll()

        If IO.File.Exists(ScriptPath + "Sparkle v1.3 2020.sls") = True Then
            Script = IO.File.ReadAllText(ScriptPath + "Sparkle v1.3 2020.sls")
        End If
        MsgBox(Script)

        SC = 0 : DC = 0 : PC = 0 : FC = 0

        NewD = False

        ConvertScriptToNodes(ScriptNode, Script)

        TV.ExpandAll()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddNewEntryNode(Parent As TreeNode)
        On Error GoTo Err

        Parent.Nodes.Add(sAddEntry, "[Add New Entry]")
        Parent.Nodes(sAddEntry).ForeColor = Color.Navy

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddScriptNode(Name As String, Path As String, Optional Tag As Integer = 0, Optional NodeColor As Color = Nothing, Optional NodeFnt As Font = Nothing)
        On Error GoTo Err

        ScriptNode.Nodes.Add(Name, "Script:" + vbTab + Path)
        ScriptNode.Nodes(Name).Tag = Tag

        ScriptNode.Nodes(Name).ForeColor = If(NodeColor = Nothing, Color.Black, NodeColor)

        If NodeFnt IsNot Nothing Then
            ScriptNode.Nodes(Name).NodeFont = NodeFnt
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ConvertScriptToNodes(SN As TreeNode, S As String, Optional SubScript As Boolean = False)
        On Error GoTo Err

        If S = "" Then Exit Sub

        Dim Lines() As String = S.Split(Environment.NewLine)

        If SubScript = False Then
            If Lines(0) <> ScriptHeader Then
                MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
                Exit Sub
            End If
        End If

        'Empy line (new part) will be a single $0A charachter, otherwise, line string will not contain vbNewLine
        Dim DiskNode As TreeNode = TV.SelectedNode

        For I As Integer = If(SubScript = False, 1, 0) To Lines.Count - 1

            Lines(I) = Lines(I).TrimStart(vbLf)

            If InStr(Lines(I), vbTab) = 0 Then
                ScriptEntryType = Lines(I)
            Else
                ScriptEntryType = Strings.Left(Lines(I), InStr(Lines(I), vbTab) - 1)
                ScriptEntry = Strings.Right(Lines(I), Lines(I).Length - InStr(Lines(I), vbTab)).TrimStart(vbTab)
            End If

            SplitEntry()

            Select Case LCase(ScriptEntryType)
                Case ""     'New Part
                    NewPart = True
                Case "path:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    If InStr(ScriptEntryArray(0), ":") = 0 Then
                        ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    UpdateNode(DiskNode.Nodes(sDiskPath + DC.ToString), sDiskPath + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt) ', tDiskPath)
                    NewPart = True
                Case "header:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    UpdateNode(DiskNode.Nodes(sDiskHeader + DC.ToString), sDiskHeader + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt) ', tDiskHeader)
                    NewPart = True
                Case "id:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    UpdateNode(DiskNode.Nodes(sDiskID + DC.ToString), sDiskID + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt) ', tDiskID)
                    NewPart = True
                Case "name:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    UpdateNode(DiskNode.Nodes(sDemoName + DC.ToString), sDemoName + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt) ', tDemoName)
                    NewPart = True
                Case "start:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    UpdateNode(DiskNode.Nodes(sDemoStart + DC.ToString), sDemoStart + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, Color.DarkGreen, Fnt) ', tDemoStart)
                    NewPart = True
                Case "dirart:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    If ScriptEntryArray(0) <> "" Then
                        If InStr(ScriptEntryArray(0), ":") = 0 Then
                            ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
                        End If
                    End If
                    UpdateNode(DiskNode.Nodes(sDirArt + DC.ToString), sDirArt + ScriptEntryArray(0), DiskNode.Tag, Color.DarkGreen, Fnt)
                    NewPart = True
                Case "packer:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    Dim Fnt As New Font("Consolas", 10)
                    UpdateNode(DiskNode.Nodes(sPacker + DC.ToString), sPacker + LCase(ScriptEntryArray(0)), DiskNode.Tag, Color.DarkGreen, Fnt)
                    Select Case LCase(ScriptEntryArray(0))
                        Case "faster"
                            Packer = 1
                        Case Else   '"better"
                            Packer = 2
                    End Select
                    NewPart = True
                Case "zp:"
                    If NewD = False Then
                        NewD = True
                        DiskNode = NewDiskToTree(SN)
                    End If
                    If CurrentDisk = 1 Then 'ZP can only be set from the first disk
                        'Correct length
                        If ScriptEntryArray(0).Length < 2 Then
                            ScriptEntryArray(0) = Strings.Left("02", 2 - ScriptEntryArray(0).Length) + ScriptEntryArray(0)
                        ElseIf ScriptEntryArray(0).Length > 2 Then
                            ScriptEntryArray(0) = Strings.Right(ScriptEntryArray(0), 2)
                        End If
                        Dim Fnt As New Font("Consolas", 10)
                        UpdateNode(DiskNode.Nodes(sZP + DC.ToString), sZP + "$" + LCase(ScriptEntryArray(0)), DiskNode.Tag, Color.DarkGreen, Fnt)
                    End If
                    NewPart = True
                Case "list:", "script:"
                    If InStr(ScriptEntryArray(0), ":") = 0 Then
                        ScriptEntryArray(0) = ScriptPath + ScriptEntryArray(0)
                    End If

                    If IO.File.Exists(ScriptEntryArray(0)) = True Then
                        AddScriptNode(SN, ScriptEntryArray(0))
                        ConvertScriptToNodes(ScriptNode.Nodes("S" + SC.ToString), IO.File.ReadAllText(ScriptEntryArray(0)), SubScript:=True)
                    Else
                        MsgBox(ScriptEntryArray(0) + " does not exist")
                    End If

                    Dim NewS As String

                    'InsertList(ScriptEntryArray(0))
                Case "file:"
                    NewD = False
                    AddFileFromScript(SN)
                    NewPart = False
                Case "new disk"
                    'DiskNode = NewDiskToTree(DiskNode)
                    ''Update last part size and disk size before starting new disk node
                    ''If DiskNode.Nodes.Count > 7 Then UpdatePartSize(DiskNode)   'First 6 nodes are disk info, last one is AddPart node
                    'UpdatePartSize(DiskNode)
                    'CurrentDisk = DiskCnt
                    'DiskNode.Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk - 1).ToString + " block" + IIf(DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " used, " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " free]"

                    'GoTo NewDisk
                Case Else
                    'Figure out what to do with comments here...
            End Select
        Next

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
    Private Function NewDiskToTree(DiskNode As TreeNode) As TreeNode

        ''Update last part size and disk size before starting new disk node
        ''If DiskNode.Nodes.Count > 7 Then UpdatePartSize(DiskNode)   'First 6 nodes are disk info, last one is AddPart node

        'UpdatePartSize(DiskNode)

        'CurrentDisk = DiskCnt
        'DiskNode.Text = "[Disk " + CurrentDisk.ToString + ": " + DiskSizeA(CurrentDisk - 1).ToString + " block" + IIf(DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " used, " + (664 - DiskSizeA(CurrentDisk - 1)).ToString + " block" + IIf(664 - DiskSizeA(CurrentDisk - 1) <> 1, "s", "") + " free]"

        TV.SelectedNode = ScriptNode.Nodes(sAddEntry)

        'Reset buffer and other disk variables here
        ResetDiskVariables()

        BlankDiskStructure()

        NewDiskToTree = TV.SelectedNode

    End Function

    Private Sub BlankDiskStructure()
        On Error GoTo Err

        TV.Enabled = False

        Dim N As TreeNode = ScriptNode.Nodes(sAddEntry)    'TV.SelectedNode

        DC += 1

        If DiskSizeA.Count < DC Then
            ReDim Preserve DiskSizeA(DC - 1)
        End If

        CurrentDisk = DC

        TV.BeginUpdate()

        With N
            .Text = "[Disk " + DC.ToString + "]"
            .Name = "D" + DC.ToString
            .ForeColor = Color.DarkRed
            '.BackColor = Color.LightPink
            .Tag = DC * &H1000000
        End With

        Dim Fnt As New Font("Consolas", 10)

        AddNode(N, sDiskPath + DC.ToString, sDiskPath + "C:\demo.d64", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDiskHeader + DC.ToString, sDiskHeader + "demo disk " + Year(Now).ToString, N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDiskID + DC.ToString, sDiskID + "sprkl", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDemoName + DC.ToString, sDemoName + "demo", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDemoStart + DC.ToString, sDemoStart + "$", N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sDirArt + DC.ToString, sDirArt, N.Tag, Color.DarkGreen, Fnt)
        AddNode(N, sPacker + DC.ToString, sPacker + IIf(My.Settings.DefaultPacker = 1, "faster", "better"), N.Tag, Color.DarkGreen, Fnt)
        If DC = 1 Then
            AddNode(N, sZP + DC.ToString, sZP + "$02", N.Tag, Color.DarkGreen, Fnt)
        End If

        AddNewEntryNode(ScriptNode)

        'AddNewPartNode(N)       '[Add new part...]

        'UpdateNewPartNode()     '->[Part 1] + [Add new file...]

        'AddNewDiskNode()        '[Add new disk...]

        TV.Enabled = True
        TV.SelectedNode = N
        TV.EndUpdate()
        TV.ExpandAll()
        TV.Focus()

        'CalcDiskNodeSize(N)

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

    Private Sub AddFileFromScript(SN As TreeNode)

        If NewPart = True Then AddPartNode(SN)

    End Sub

    Private Sub AddPartNode(SN As TreeNode)

        PC += 1

        UpdateNode(SN.Nodes(sAddEntry), "[Part " + PC.ToString + "]")
        SN.Nodes(sAddEntry).Name = "P" + PC.ToString
        SN.Nodes("P" + PC.ToString).ForeColor = Color.DarkMagenta
        AddNewEntryNode(ScriptNode)

        NewPart = False

    End Sub

    Private Sub AddScriptNode(SN As TreeNode, S As String)

        SC += 1

        UpdateNode(SN.Nodes(sAddEntry), "Script: " + S)
        SN.Nodes(sAddEntry).Name = "S" + SC.ToString
        With SN.Nodes("S" + SC.ToString)
            '.BackColor = Color.LightGray
            .ForeColor = Color.Black
        End With
        'NewScriptNode = ScriptNode.Nodes("S" + SC.ToString)
        AddNewEntryNode(SN.Nodes("S" + SC.ToString))
        If ScriptNode.Nodes(sAddEntry) Is Nothing Then
            AddNewEntryNode(ScriptNode)
        End If
    End Sub


End Class