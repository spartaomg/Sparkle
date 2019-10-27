Friend Module ModDisk

    Public Drive() As Byte ' = My.Resources.SparkleDrive

    Public S2 As Boolean = False

    Public PRGFull() As String
    Public PRGShort() As String
    Public PRGBlock() As Byte
    Public PRGTrack() As Byte
    Public PRGSector() As Byte
    Public PRGCnt As Integer
    Public BlocksFree As Integer = 664

    Public Disk(174847), NextTrack, NextSector As Byte   'Next Empty Track and Sector
    Public MaxSector As Byte = 18, LastSector, BlockCnt, Prg() As Byte

    Public FileUnderIO As Boolean = False
    Public IOBit As Byte

    Public Track(35), CT, CS, CP, BufferCt As Integer
    Public TestDisk As Boolean = False
    Public StartTrack As Byte = 1
    Public StartSector As Byte = 0

    Public ByteSt(), BitSt(), Buffer(255), BitPos, LastByte, AdLo, AdHi, MaxType As Byte
    Public Match, MaxBit, MatchSave(), MaxSave, PrgLen, Distant As Integer
    Public MatchOffset(), MatchCnt, RLECnt, MatchLen(), MaxOffset, MaxLen, LitCnt, Bits, BuffAdd, PrgAdd As Integer
    Public DistAd(), DistLen(), DistSave(), DistCnt, DistBase As Integer
    Public DtPos, CmPos, CmLast, DtLen, MatchStart, BitCt As Integer
    Public ByteCt As Integer    'Points at the next empty byte in buffer

    'Public Compression As String

    Public Script As String
    Public ScriptHeader As String = "[Sparkle Loader Script]"
    Public ScriptName As String
    Public DiskNo As Integer
    'Public D64Path As String

    '-----THESE WILL BE REMOVED-----
    Public D64Name As String '= My.Computer.FileSystem.SpecialDirectories.MyDocuments
    'Public DiskHeader As String = "OMG " + Year(Now).ToString
    '-------------------------------

    Public DiskHeader As String = "demo disk " + Year(Now).ToString
    Public DiskID As String = "sprkl"
    'Public DiskComp As String
    Public DemoName As String = "demo"
    Public DemoStart As String
    Public Music, MusicInit, MusicPlay As String
    Public AddLoader As Boolean = True
    Public SystemFile As Boolean = False
    Public NextFileCnt As Byte = 0
    Public NextID As Byte = 0
    Public PRGList As String
    Public FileChanged As Boolean = False

    Public DiskCnt As Integer = -1
    Public PartCnt As Integer = -1
    Public FileCnt As Integer = -1
    Public CurrentDisk As Integer = -1
    Public CurrentPart As Integer = -1
    Public CurrentFile As Integer = -1

    Public D64NameA(), DiskHeaderA(), DiskIDA(), DemoNameA(), DemoStartA() As String
    Public FilesA(), FileAddrA(), FileOffsA(), FileLenA() As String
    Public Prgs As New List(Of Byte())
    Public FileIOA() As Boolean

    Public DiskNoA(), DFDiskNoA(), DFPartNoA(), DiskPartCntA(), DiskFileCntA() As Integer
    Public FilesInPartA() As Integer
    Public PDiskNoA(), PSizeA() As Integer
    Public FDiskNoA(), FPartNoA(), FSizeA() As Integer
    Public TotalParts As Integer = 0
    Public NewFile As String

    Public DiskBlockSize() As Integer
    Public PartBlockSize() As Integer
    Public FileBlockSize() As Integer
    Public FBSDisk() As Integer
    Public MusicBlockSize As Integer = 0

    Public bBuildDisk As Boolean = False

    Dim SS, SE As Integer
    'Dim EndOfSection As Boolean = False
    Dim NewPart As Boolean = False
    Dim ScriptEntryType As String = ""
    Dim ScriptEntry As String = ""
    Dim ScriptEntryArray() As String
    Dim LastNonEmpty As Integer = -1

    Public LC(), NM, FM, LM As Integer
    Public SM1 As Integer = 0
    Public SM2 As Integer = 0
    Public OverMaxLit As Boolean = False

    Dim FirstFileOfDisk As Boolean = False
    Dim FirstFileStart As String = ""

    Public TabT(663), TabS(663) As Byte

    Public Sub SetLastSector()
        On Error GoTo Err

        Select Case CT
            Case 1 To 17
                MaxSector = 20
                LastSector = 17
            Case 18 To 24
                MaxSector = 18
                LastSector = 15
            Case 25 To 30
                MaxSector = 17
                LastSector = 15
            Case 31 To 35
                MaxSector = 16
                LastSector = 13
        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function ConvertScriptToArray() As Boolean
        On Error GoTo Err

        ConvertScriptToArray = False

        SS = 1 : SE = 1

        FindNextScriptEntry()

        If ScriptEntry <> ScriptHeader Then
            MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
            Exit Function
        End If

        ResetArrays()
        AddNewDiskToArray()

FindNext:
        FindNextScriptEntry()
        FindEntryType()

        If SE < Script.Length Then GoTo FindNext

        DiskPartCntA(CurrentDisk) = TotalParts
        'MsgBox("Disk(" + CurrentDisk.ToString + "): " + DiskPartCntA(CurrentDisk).ToString + " parts")
        ConvertScriptToArray = True

        Exit Function

Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        ConvertScriptToArray = False

    End Function

    Public Sub ResetArrays()
        On Error GoTo Err

        DiskCnt = -1

        ReDim DiskNoA(DiskCnt), DiskPartCntA(DiskCnt), DiskFileCntA(DiskCnt)
        ReDim D64NameA(DiskCnt), DiskHeaderA(DiskCnt), DiskIDA(DiskCnt), DemoNameA(DiskCnt), DemoStartA(DiskCnt)
        ReDim DiskBlockSize(DiskCnt)

        PartCnt = -1
        ReDim PDiskNoA(PartCnt), PSizeA(PartCnt), PartBlockSize(PartCnt), FilesInPartA(PartCnt)

        FileCnt = -1

        ReDim FilesA(FileCnt), DFDiskNoA(FileCnt), DFPartNoA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt)
        ReDim FDiskNoA(FileCnt), FPartNoA(FileCnt), FSizeA(FileCnt)
        ReDim FileBlockSize(FileCnt), FBSDisk(FileCnt)

        TotalParts = 0

        MusicBlockSize = 0

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub AddNewDiskToArray()
        On Error GoTo Err

        DiskCnt += 1
        ReDim Preserve DiskNoA(DiskCnt), DiskPartCntA(DiskCnt), DiskFileCntA(DiskCnt)
        ReDim Preserve D64NameA(DiskCnt), DiskHeaderA(DiskCnt), DiskIDA(DiskCnt), DemoNameA(DiskCnt), DemoStartA(DiskCnt)
        'ReDim Preserve MusicA(DiskCnt), MusicInitA(DiskCnt), MusicPlayA(DiskCnt)
        ReDim Preserve DiskBlockSize(DiskCnt)

        DiskNoA(DiskCnt) = DiskCnt
        CurrentDisk = DiskCnt
        TotalParts = 0

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub AddNewPartToArray(DiskIndex As Integer)
        On Error GoTo Err

        PartCnt += 1
        ReDim Preserve FilesInPartA(PartCnt), PDiskNoA(PartCnt), PSizeA(PartCnt), PartBlockSize(PartCnt)

        PDiskNoA(PartCnt) = DiskIndex
        CurrentPart = PartCnt
        'TotalParts += 1

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub AddNewFileToArray(DiskIndex As Integer)
        On Error GoTo Err
        'If MsgBox("Add " + NewFile + " to Disk " + DiskIndex.ToString + ", Part " + CurrentPart.ToString + "?", vbOKCancel + vbCancel, "Add new file?") = vbCancel Then Exit Sub

        FileCnt += 1
        ReDim Preserve FilesA(FileCnt), DFDiskNoA(FileCnt), DFPartNoA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt)
        ReDim Preserve FDiskNoA(FileCnt), FPartNoA(FileCnt), FSizeA(FileCnt)
        ReDim Preserve FileBlockSize(FileCnt), FBSDisk(FileCnt)

        DFDiskNoA(FileCnt) = DiskIndex
        DFPartNoA(FileCnt) = CurrentPart
        FilesInPartA(CurrentPart) += 1            'Number
        CurrentFile = FileCnt
        DiskFileCntA(DiskIndex) += 1

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FindNextScriptEntry()
        On Error GoTo Err

        If Mid(Script, SS, 1) = Chr(13) Then                'Check if this is an empty line indicating a new section
            NewPart = True
            SS += 2                                         'Skip EOL bytes
            SE = SS + 1
        Else
            NewPart = False
        End If
NextChar:
        If Mid(Script, SE, 1) <> Chr(13) Then               'Look for EOL
            SE += 1                                         'Not EOL
            If SE <= Script.Length Then                     'Go to next char if we haven't reached the end of the script
                GoTo NextChar
            Else
                ScriptEntry = Strings.Mid(Script, SS, SE - SS)    'Reached end of script, finish this entry
            End If
        Else                                                'Found EOL
            ScriptEntry = Strings.Mid(Script, SS, SE - SS)        'Finish this entry
            SS = SE + 2                                     'Skip EOL bytes
            SE = SS + 1
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FindEntryType()
        On Error GoTo Err

        If Strings.InStr(ScriptEntry, vbTab) = 0 Then
            ScriptEntryType = ScriptEntry
        Else
            ScriptEntryType = Strings.Left(ScriptEntry, Strings.InStr(ScriptEntry, vbTab) - 1)
            ScriptEntry = Strings.Right(ScriptEntry, ScriptEntry.Length - Strings.InStr(ScriptEntry, vbTab))
        End If

        SplitEntry()

        Select Case ScriptEntryType
            Case "Path:"
                D64NameA(CurrentDisk) = ScriptEntryArray(0)
            Case "Header:"
                DiskHeaderA(CurrentDisk) = ScriptEntryArray(0)
            Case "ID:"
                DiskIDA(CurrentDisk) = ScriptEntryArray(0)
            Case "Name:"
                DemoNameA(CurrentDisk) = ScriptEntryArray(0)
            Case "Start:"
                DemoStartA(CurrentDisk) = ScriptEntryArray(0)
            Case "File:"
                AddFileToArray()
                FilesA(FileCnt) = ScriptEntryArray(0)                                 'New File's Path
                FPartNoA(FileCnt) = PartCnt
                If ScriptEntryArray.Count > 1 Then FileAddrA(FileCnt) = ScriptEntryArray(1) 'New File's load address
                If ScriptEntryArray.Count > 2 Then FileOffsA(FileCnt) = ScriptEntryArray(2) 'New File's offset within file
                If ScriptEntryArray.Count = 4 Then FileLenA(FileCnt) = ScriptEntryArray(3)  'New File's length in bytes within file
            Case "New Disk"
                DiskPartCntA(CurrentDisk) = TotalParts
                'MsgBox("Disk(" + CurrentDisk.ToString + "): " + DiskPartCntA(CurrentDisk).ToString + " parts")
                AddNewDiskToArray()
            Case Else
        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddFileToArray()
        On Error GoTo Err

        If NewPart = True Then
            AddNewPartToArray(CurrentDisk)
            TotalParts += 1
        End If

        AddNewFileToArray(CurrentDisk)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SplitEntry()
        On Error GoTo Err

        LastNonEmpty = -1

        ReDim ScriptEntryArray(LastNonEmpty)

        ScriptEntryArray = Split(ScriptEntry, vbTab)

        For I As Integer = 0 To ScriptEntryArray.Length - 1
            If ScriptEntryArray(I) <> "" Then
                LastNonEmpty += 1
                ScriptEntryArray(LastNonEmpty) = ScriptEntryArray(I)
            End If
        Next

        If LastNonEmpty > -1 Then
            ReDim Preserve ScriptEntryArray(LastNonEmpty)
        Else
            ReDim ScriptEntryArray(0)
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub
    Public Sub NewDisk()
        On Error GoTo Err

        ReDim Disk(174847)
        Dim B As Byte

        BlocksFree = 664

        'D64Name = ""
        'FileChanged = False
        'StatusFileName(D64Name)

        Dim Cnt As Integer

        CP = Track(18)

        Disk(CP) = &H12         'Track#18
        Disk(CP + 1) = &H1      'Sector#1
        Disk(CP + 2) = &H41     '"A"

        For Cnt = &H90 To &HAA  'Name, ID, DOS type
            Disk(CP + Cnt) = &HA0
        Next

        For Cnt = 1 To Strings.Len(DiskHeader)
            B = Asc(Mid(DiskHeader, Cnt, 1))
            If B > &H5F Then B -= &H20
            Disk(CP + &H8F + Cnt) = B
        Next

        For Cnt = 1 To Len(DiskID)                  'SPRKL
            B = Asc(Mid(DiskID, Cnt, 1))
            If B > &H5F Then B -= &H20
            Disk(CP + &HA1 + Cnt) = B
        Next

        For Cnt = 4 To 36 * 4 - 1
            Disk(CP + Cnt) = 255
        Next

        For Cnt = 1 To 17
            Disk(CP + (Cnt * 4) + 0) = 21
            Disk(CP + (Cnt * 4) + 3) = 31
        Next

        For Cnt = 18 To 24
            Disk(CP + (Cnt * 4) + 0) = 19
            Disk(CP + (Cnt * 4) + 3) = 7
        Next

        For Cnt = 25 To 30
            Disk(CP + (Cnt * 4) + 0) = 18
            Disk(CP + (Cnt * 4) + 3) = 3
        Next

        For Cnt = 31 To 35
            Disk(CP + (Cnt * 4) + 0) = 17
            Disk(CP + (Cnt * 4) + 3) = 1
        Next

        Disk(CP + (18 * 4) + 0) = 17
        Disk(CP + (18 * 4) + 1) = 252

        CT = 18
        CS = 1

        SetLastSector()

        CP = Track(CT) + 256 * CS
        Disk(CP + 1) = 255

        NextTrack = 1 : NextSector = 1

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function BuildDisk(DiskIndex As Integer) As Boolean
        On Error GoTo Err

        BuildDisk = True

        'Dim TotalBC As Integer
        Dim FC, PC As Integer

        StartTrack = 1 : StartSector = 0

        NextTrack = StartTrack
        NextSector = StartSector

        BlockCnt = 0

        ReDim ByteSt(0)
        ResetBuffer()       'This is probably not needed here

        SM1 = 0
        SM2 = 0
        LM = 0
        FirstFileOfDisk = True
        PC = -1
        'TotalBC = 0
        If FileCnt > -1 Then
            For FC = 0 To FileCnt
                If DFDiskNoA(FC) = DiskIndex Then   'Is this file on the current disk?
                    If DFPartNoA(FC) <> PC Then     'Is this te first file of a new part?
                        'Yes, first file of new part
                        If PC > -1 Then
                            'Close last buffer
                            CloseLastBuff()
                            'Add Block Count to part
                            ByteSt(255) = BlockCnt
                            'Add compressed part to disk
                            AddCompressedPart(NextTrack, NextSector)
                            'Calculate total block count here
                            'TotalBC += BlockCnt
                        End If
                        'Start next Part
                        PC = DFPartNoA(FC)
                        'Reset Block Counter
                        BlockCnt = 0
                        'Reset Buffer for BlockCnt=0 (BytCt=254 is NOT enough!!!), needed for first part here
                        ResetBuffer()
                    End If
                    If FirstFileOfDisk = True Then
                        FirstFileStart = FileAddrA(FC)
                        FirstFileOfDisk = False
                    End If
                    'Add file to part
                    LZ(FilesA(FC), FileAddrA(FC), FileOffsA(FC), FileLenA(FC))
                    'MsgBox(FilesA(FC) + vbNewLine + DFPartNoA(FC).ToString + vbNewLine + BlockCnt.ToString)
                End If
            Next
            'Close last buffer of last part
            CloseLastBuff()
            'Add Block count to last part
            ByteSt(255) = BlockCnt
            'Add last compressed part to disk
            AddCompressedPart(NextTrack, NextSector)
            'Calculate total block count
            'TotalBC += BlockCnt
        End If

        'MsgBox("Block Count: " + TotalBC.ToString + " blocks" + vbNewLine +
        '       "SM1:" + vbTab + SM1.ToString + vbNewLine +
        '       "SM2:" + vbTab + SM2.ToString + vbNewLine +
        '       "LM:" + vbTab + LM.ToString)

        SystemFile = True
        '               This Disk's ID    Part count on Disk      Next Disk's ID or 0 if this is the last disk     
        If InjectDriveCode(DiskIndex + 1, DiskPartCntA(CurrentDisk), IIf(DiskIndex < DiskCnt, DiskIndex + 2, 0)) = False Then GoTo NoDisk

        If InjectLoader(DiskIndex, 18, 5, 6) = False Then GoTo NoDisk

        SystemFile = False

        If UpdateBAM(DiskIndex) = False Then GoTo NoDisk

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        BuildDisk = False

    End Function

    Public Sub AddCompressedPart(Optional T As Byte = 1, Optional S As Byte = 0)
        On Error GoTo Err

        If BlocksFree < BlockCnt Then
            MsgBox("Unable to add part to disk", vbOKOnly, "Not enough free space on disk")
            Exit Sub
        End If

        CT = T
        CS = S

        If SectorOK(CT, CS) = False Then
            FindNextFreeSector()
        End If

        For I = 0 To BlockCnt - 1
            For J = 0 To 255
                'Disk(Track(CT) + 256 * CS + J) = Prg(I * 256 + J)
                Disk(Track(CT) + 256 * CS + J) = ByteSt(I * 256 + J)
            Next

            DeleteBit(CT, CS, True)

            If AddInterleave() = False Then   'No free sectors left on disk
                If I = BlockCnt - 1 Then
                    'Last block of file in last block of disk
                    'So we go to 18:00 to add Next Side Info
                    CT = 18
                    CS = 0
                Else
                    'Reached last sector, but file copy is not complete (this should not happen)
                    MsgBox("Unable to complete disk building!", vbOKOnly, "Disk is full")
                    Exit For
                End If
            End If   'Go to next free sector with Interleave IL
        Next

        NextTrack = CT
        NextSector = CS

        'If AddToDir = True Then
        'AddFileToDir(PrgName)
        'End If

        ReDim ByteSt(0)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function SectorOK(T As Byte, S As Byte) As Boolean
        On Error GoTo Err

        Dim BP As Integer   'BAM Position for Bit Change
        Dim BB As Integer   'BAM Bit

        BP = Track(18) + T * 4 + 1 + Int(S / 8)
        BB = 2 ^ (S Mod 8)    '=0-7
        If (Disk(BP) And BB) = 0 Then   'Block is already used
            SectorOK = False
        Else
            SectorOK = True                 'Block is unused
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        SectorOK = False

    End Function

    Private Sub FindNextFreeSector()
        On Error GoTo Err

CheckB:
        If SectorOK(CT, CS) = False Then
            CS += 1
            Select Case CT
                Case 1 To 17
                    If CS = 21 Then CS = 0
                Case 18 To 24
                    If CS = 19 Then CS = 0
                Case 25 To 30
                    If CS = 18 Then CS = 0
                Case 31 To 35
                    If CS = 17 Then CS = 0
            End Select
            GoTo CheckB
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub DeleteBit(T As Byte, S As Byte, Optional UpdateBlocksFree As Boolean = True)
        On Error GoTo Err

        Dim BP As Integer   'BAM Position for Bit Change
        Dim BB As Integer   'BAM Bit

        BP = Track(18) + T * 4 + 1 + Int(S / 8)
        BB = 255 - 2 ^ (S Mod 8)    '=0-7
        Disk(BP) = Disk(BP) And BB

        BP = Track(18) + T * 4
        Disk(BP) -= 1

        If UpdateBlocksFree = True Then BlocksFree -= 1

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function AddInterleave(Optional IL As Byte = 5) As Boolean
        On Error GoTo Err

        AddInterleave = True

        If TrackIsFull(CT) = True Then 'If this track is full, go to next and check again
            If (CT = 35) Or (CT = 18) Then         'Reached max track No, disk is full
                AddInterleave = False
                Exit Function
            End If
            CalcNextSector(IL)
            CT += 1

            If SystemFile = False Then
                If CT = 18 Then
                    CT = 19
                    CS = 3  'Need to skip 2 sectors while skipping Track 18 (the disk keeps spinning)
                End If
            End If
            'First sector in new track will be #1 and NOT #0!!!
            'CS = StartSector
        Else
            CalcNextSector(IL)
        End If
        FindNextFreeSector()

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        AddInterleave = False

    End Function

    Private Function TrackIsFull(T As Byte) As Boolean
        On Error GoTo Err

        If Disk(Track(18) + T * 4) = 0 Then
            TrackIsFull = True
        Else
            TrackIsFull = False
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Sub CalcNextSector(Optional IL As Byte = 5)
        On Error GoTo Err

        Select Case CT

            Case 1 To 17    '21 sectors, 0-20

                If CS > 20 Then
                    CS -= 21
                    If CS > 0 Then CS -= 1
                End If
                CS += 4     'IL=4 always
                If CS > 20 Then
                    CS -= 21
                    If CS > 0 Then CS -= 1
                End If

            Case 18         'Handle Dir Track separately
                If CS > 18 Then CS -= 19
                CS += IL
                If CS > 18 Then CS -= 19

            Case 19 To 24   '19 sectors, 0-18
                If CS > 18 Then CS -= 19
                CS += 3     'IL=3 always
                If CS > 18 Then CS -= 19

            Case 25 To 30   '18 sectors, 0-17
                If CS > 17 Then CS -= 18
                CS += 3     'IL=3 always
                If CS > 17 Then CS -= 18

            Case 31 To 35   '17 sectors, 0-16
                If CS > 16 Then CS -= 17
                CS += 3     'IL=3 always
                If CS > 16 Then CS -= 17

        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function InjectDriveCode(idcDiskID As Byte, idcFileCnt As Byte, idcNextID As Byte) As Boolean
        On Error GoTo Err

        InjectDriveCode = True

        Dim I, Cnt As Integer

        Drive = My.Resources.SD

        Drive(802) = idcFileCnt 'Save number of parts to be loaded to ZP Tab Location $20
        Drive(803) = idcNextID  'Save Next Side ID1 to ZP Tab Location $21

        CT = 18
        CS = 2
        For Cnt = 0 To 4        '5 blocks to be saved: 18:02, 18:06, 18:10, 18:14, 18:18
            For I = 0 To 255
                If Drive.Length > Cnt * 256 + I + 2 Then
                    Disk(Track(CT) + CS * 256 + I) = Drive(Cnt * 256 + I + 2)
                End If
            Next
            DeleteBit(CT, CS, False)
            CS += 4
        Next

        'Next Side Info on last 3 bytes of BAM!!!
        Disk(Track(18) + 0 * 256 + 255) = idcDiskID
        Disk(Track(18) + 0 * 256 + 254) = idcFileCnt
        Disk(Track(18) + 0 * 256 + 253) = idcNextID

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        InjectDriveCode = False

    End Function

    Public Function InjectLoader(DiskIndex As Integer, T As Byte, S As Byte, IL As Byte) As Boolean
        On Error GoTo Err

        InjectLoader = True

        Dim B, I, Cnt, W As Integer
        Dim ST, SS, A, L, AdLo, AdHi As Byte
        Dim Loader() As Byte
        Dim DN As String

        If DiskIndex > -1 Then
            B = ConvertHexStringToNumber(DemoStartA(DiskIndex))
        Else
            B = ConvertHexStringToNumber(DemoStart)
        End If

        If B = 0 Then B = ConvertHexStringToNumber(FirstFileStart)

        If B = 0 Then
            MsgBox("Unable to build demo disk." + vbNewLine + vbNewLine + "Missing start address", vbOKOnly)
            InjectLoader = False
            Exit Function
        End If

        AdLo = (B - 1) Mod 256
        AdHi = Int((B - 1) / 256)

        Loader = My.Resources.SL

        For I = 0 To Loader.Length - 3       'Find JMP $01f0 instruction (JMP AltLoad)
            If (Loader(I) = &H4C) And (Loader(I + 1) = &HF0) And (Loader(I + 2) = &H1) Then
                Loader(I - 1) = AdLo       'Lo Byte return address at the end of Loader
                Loader(I - 3) = AdHi       'Hi Byte return address at the end of Loader
                Exit For
            End If
        Next

        'Number of blocks in Loader
        L = Int(Loader.Length / 254)
        If (Loader.Length) Mod 254 <> 0 Then
            L += 1
        End If

        CT = T
        CS = S

        For I = 0 To L - 1
            ST = CT
            SS = CS
            For Cnt = 0 To 253
                If I * 254 + Cnt < Loader.Length Then
                    Disk(Track(CT) + CS * 256 + 2 + Cnt) = Loader(I * 254 + Cnt)
                End If
            Next
            DeleteBit(CT, CS, True)

            AddInterleave(IL)   'Go to next free sector with Interleave IL
            If I < L - 1 Then
                Disk(Track(ST) + SS * 256 + 0) = CT
                Disk(Track(ST) + SS * 256 + 1) = CS
            Else
                Disk(Track(ST) + SS * 256 + 0) = 0
                If Loader.Length Mod 254 = 0 Then
                    Disk(Track(ST) + SS * 256 + 1) = 254 + 1
                Else
                    Disk(Track(ST) + SS * 256 + 1) = ((Loader.Length) Mod 254) + 1
                End If
            End If

        Next

        CT = 18 : CS = 1
        Cnt = Track(CT) + CS * 256
SeekNewDirBlock:
        If Disk(Cnt) <> 0 Then
            Cnt = Track(Disk(Cnt)) + Disk(Cnt + 1) * 256
            GoTo SeekNewDirBlock
        Else
            B = 2
SeekNewEntry:
            If Disk(Cnt + B) = &H0 Then
                Disk(Cnt + B) = &H82
                Disk(Cnt + B + 1) = T
                Disk(Cnt + B + 2) = S
                For W = 0 To 15
                    Disk(Cnt + B + 3 + W) = &HA0
                Next

                If DiskIndex > -1 Then
                    DN = If(DemoNameA(DiskIndex) <> "", DemoNameA(DiskIndex), "demo")
                Else
                    DN = If(DemoName <> "", DemoName, "demo")
                End If

                For W = 0 To Len(DN) - 1
                    A = Asc(Mid(DN, W + 1, 1))
                    If A > &H5F Then A -= &H20
                    Disk(Cnt + B + 3 + W) = A
                Next
                Disk(Cnt + B + &H1C) = L
                Else
                    B += 32
                If B < 256 Then
                    GoTo SeekNewEntry
                Else
                    CS += 4
                    If CS > 18 Then S -= 18
                    Disk(Cnt) = CT
                    Disk(Cnt + 1) = CS
                    Cnt = Track(CT) + CS * 256
                    Disk(Cnt) = 0
                    Disk(Cnt + 1) = 255
                    GoTo SeekNewDirBlock
                End If
            End If
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        InjectLoader = False

    End Function

    Public Function ConvertHexStringToNumber(HexString As String) As Integer
        On Error GoTo Err

        ConvertHexStringToNumber = 0

        If HexString = "" Then Exit Function

        Dim B, C, I As Integer
        B = 0
        For I = 1 To Len(HexString)
            C = Asc(Strings.Mid(HexString, I, 1))
            C -= 48
            If C > 9 Then
                C -= 7
            End If
            If C > 15 Then
                C -= 32
            End If
            B = B * 16 + C
        Next

        ConvertHexStringToNumber = B

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        ConvertHexStringToNumber = 0

    End Function

    Public Function ConvertNumberToHexString(HLo As Byte, Optional HHi As Byte = 0) As String
        On Error GoTo Err

        Dim S As String = ""
        Dim B As Byte

        'S = Hex(HLo + HHi * 256)

        'If Len(S) < 4 Then
        'S = Strings.Left("0000", 4 - Len(S)) + S
        'End If

        B = Int(HHi / 16)

        If B < 10 Then
            S += Strings.Chr(48 + B)
        Else
            S += Strings.Chr(87 + B)
        End If

        B = HHi Mod 16

        If B < 10 Then
            S += Strings.Chr(48 + B)
        Else
            S += Strings.Chr(87 + B)
        End If

        B = Int(HLo / 16)
        If B < 10 Then
            S += Strings.Chr(48 + B)
        Else
            S += Strings.Chr(87 + B)
        End If

        B = HLo Mod 16

        If B < 10 Then
            S += Strings.Chr(48 + B)
        Else
            S += Strings.Chr(87 + B)
        End If

        ConvertNumberToHexString = S

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        ConvertNumberToHexString = ""

    End Function

    Public Function UpdateBAM(DiskIndex As Integer) As Boolean
        On Error GoTo Err

        UpdateBAM = True

        Dim Cnt As Integer
        Dim B As Byte
        CP = Track(18)

        For Cnt = &H90 To &HAA  'Name, ID, DOS type
            Disk(CP + Cnt) = &HA0
        Next

        If DiskHeaderA(DiskIndex) = "" Then DiskHeaderA(DiskIndex) = "demo disk" + IIf(DiskCnt > 0, " " + (DiskIndex + 1).ToString, "")

        For Cnt = 1 To Strings.Len(DiskHeaderA(DiskIndex))
            B = Asc(Strings.Mid(DiskHeaderA(DiskIndex), Cnt, 1))
            If B > &H5F Then B -= &H20
            Disk(CP + &H8F + Cnt) = B
        Next

        If DiskIDA(DiskIndex) <> "" Then
            For Cnt = 1 To Strings.Len(DiskIDA(DiskIndex))
                B = Asc(Strings.Mid(DiskIDA(DiskIndex), Cnt, 1))
                If B > &H5F Then B -= &H20
                Disk(CP + &HA1 + Cnt) = B
            Next
        Else
            For Cnt = 1 To Strings.Len(DiskID)
                B = Asc(Strings.Mid(DiskID, Cnt, 1))
                If B > &H5F Then B -= &H20
                Disk(CP + &HA1 + Cnt) = B
            Next
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        UpdateBAM = False

    End Function

    Public Function ConvertArrayToScript() As String
        On Error GoTo Err

        Dim S As String
        Dim PC As Integer = -1
        S = ScriptHeader + vbNewLine + vbNewLine

        For I As Integer = 0 To DiskCnt
            S += "Path:" + vbTab + D64NameA(I) + vbNewLine +
                "Header:" + vbTab + DiskHeaderA(I) + vbNewLine +
                "Name:" + vbTab + DemoNameA(I) + vbNewLine +
                "Start:" + vbTab + DemoStartA(I)
            For J As Integer = 0 To FileCnt
                If PDiskNoA(DFPartNoA(J)) = I Then      'Is this file in a part on this disk?
                    S += vbNewLine
                    If DFPartNoA(J) <> PC Then
                        S += vbNewLine
                        PC = DFPartNoA(J)
                    End If
                    S += "File:" + vbTab + FilesA(J)
                    If FileAddrA(J) <> "" Then
                        S += vbTab + FileAddrA(J)
                        If FileOffsA(J) <> "" Then
                            S += vbTab + FileOffsA(J)
                            If FileLenA(J) <> "" Then S += vbTab + FileLenA(J)
                        End If
                    End If
                End If
            Next
            If I < DiskCnt Then
                S += vbNewLine + vbNewLine + "New Disk" + vbNewLine + vbNewLine
            End If
        Next

        ConvertArrayToScript = S

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        ConvertArrayToScript = ""

    End Function

    Public Function BuildDemoFromScript() As Boolean
        On Error GoTo Err

        BuildDemoFromScript = True

        SS = 1 : SE = 1

        'Check if this is a valid script
        FindNextScriptEntry()

        If ScriptEntry <> ScriptHeader Then
            MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
            GoTo NoDisk
        End If

        DiskCnt = -1

NewDisk:
        'Reset Disk image
        NewDisk()
        'Reset Disk Variables
        ResetDiskVariables()
        DiskCnt += 1

FindNext:

        FindNextScriptEntry()
        'Split String
        SplitScriptEntry()

        'Set disk variables and add files
        Select Case ScriptEntryType
            Case "Path:"
                D64Name = ScriptEntryArray(0)
            Case "Header:"
                DiskHeader = ScriptEntryArray(0)
            Case "ID:"
                DiskID = ScriptEntryArray(0)
            Case "Name:"
                DemoName = ScriptEntryArray(0)
            Case "Start:"
                DemoStart = ScriptEntryArray(0)
            Case "File:"
                'Add files to part array
                AddFile()       'If new part, it will first sort files in last part then add previous part to disk

            Case "New Disk"
                'If new disk, sort and add last part, then update part count, add loader & drive code, save disk, then GoTo NewDisk
                FinishDisk(False)

                GoTo NewDisk
            Case Else
        End Select

        If SE < Script.Length Then GoTo FindNext

        'Last disk: sort and add last part, then update part count, add loader & drive code, save disk, and we are done :)
        FinishDisk(True)

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        BuildDemoFromScript = False

    End Function

    Private Sub FinishDisk(LastDisk As Boolean)
        On Error GoTo Err

        SortPart()
        AddPartToDisk()
        InjectLoader(-1, 18, 5, 6)
        InjectDriveCode(DiskCnt + 1, PartCnt, IIf(LastDisk = False, DiskCnt + 2, 0))

        SaveDisk()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SaveDisk()
        On Error GoTo Err

        If D64Name = "" Then D64Name = "Demo Disk" + IIf(DiskCnt > 0, " " + (CurrentDisk + 1).ToString, "")

        Dim SaveDLG As New SaveFileDialog With {
            .Filter = "D64 Files (*.d64)|*.d64",
            .Title = "Save D64 File As...",
            .FileName = D64Name,
            .RestoreDirectory = True
        }

        Dim R As DialogResult = SaveDLG.ShowDialog(FrmMain)

        If R = Windows.Forms.DialogResult.OK Then
            D64Name = SaveDLG.FileName
            If Strings.Right(D64Name, 4) <> ".d64" Then
                D64Name += ".d64"
            End If
            IO.File.WriteAllBytes(D64Name, Disk)
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddPartToDisk()
        On Error GoTo Err

        If Prgs.Count = 0 Then Exit Sub

        ReDim ByteSt(0)
        ResetBuffer()

        For I As Integer = 0 To Prgs.Count - 1
            LZ(Prgs(I).ToArray, FileAddrA(I),, FileLenA(I), FileIOA(I))
        Next
        CloseLastBuff()

        ByteSt(255) = BlockCnt
        AddCompressedPart(NextTrack, NextSector)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddFile()
        On Error GoTo Err

        If NewPart = True Then
            'First finish last part, if it exists
            If PartCnt > 0 Then
                'Sort files in part
                SortPart()
                AddPartToDisk()

            End If

            'Then add new part and  reset part variables
            ResetPartVariables()

        End If

        'Then add file to part
        AddFileToPart()

        'MsgBox("Part: " + PartCnt.ToString + vbNewLine + "File: " + Prgs.Count.ToString + vbNewLine + "Len: " + Hex(Prgs(FileCnt).Length)) 'FileLenA(FileCnt))

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SortPart()
        On Error GoTo Err

        Dim Change As Boolean
        Dim FAO, FLO, FAI, FLI As Integer
        Dim PO(), PI() As Byte
        Dim S As String
        Dim IO As Boolean

        'Append adjacent files
Restart:
        Change = False

        For O As Integer = 0 To Prgs.Count - 2
            FAO = ConvertHexStringToNumber(FileAddrA(O))
            FLO = ConvertHexStringToNumber(FileLenA(O))
            For I As Integer = O + 1 To Prgs.Count - 1
                FAI = ConvertHexStringToNumber(FileAddrA(I))
                FLI = ConvertHexStringToNumber(FileLenA(I))

                'MsgBox("FAO: " + Hex(FAO) + vbNewLine + "FLO: " + Hex(FLO) + vbNewLine + "FAI: " + Hex(FAI) + vbNewLine + "FLI: " + Hex(FLI))

                If FAO + FLO = FAI Then
                    'Append
                    'MsgBox("Append File" + O.ToString + " to File  " + I.ToString)
                    PO = Prgs(O)
                    PI = Prgs(I)
                    ReDim Preserve PO(FLO + FLI - 1)

                    For J As Integer = 0 To FLI - 1
                        PO(FLO + J) = PI(J)
                    Next

                    Prgs(O) = PO
                    Change = True
                ElseIf FAI + FLI = FAO Then
                    'Prepend
                    'MsgBox("Prepend File " + I.ToString + " to File " + O.ToString)
                    PO = Prgs(O)
                    PI = Prgs(I)
                    ReDim Preserve PI(FLI + FLO - 1)

                    For J As Integer = 0 To FLO - 1
                        PI(FLI + J) = PO(J)
                    Next

                    Prgs(O) = PI

                    FileAddrA(O) = FileAddrA(I)

                    Change = True
                End If

                If Change = True Then
                    FLO += FLI
                    FileLenA(O) = ConvertNumberToHexString(FLO Mod 256, Int(FLO / 256))
                    For J As Integer = I To Prgs.Count - 2
                        FileAddrA(J) = FileAddrA(J + 1)
                        FileOffsA(J) = FileOffsA(J + 1)     'this may not be needed later
                        FileLenA(J) = FileLenA(J + 1)
                    Next
                    FileCnt -= 1
                    ReDim Preserve FileAddrA(Prgs.Count - 2), FileOffsA(Prgs.Count - 2), FileLenA(Prgs.Count - 2)
                    Prgs.Remove(Prgs(I))
                    GoTo Restart
                End If
            Next
        Next

        '--------------------------------------------------------------------------------
        'Sort files by length (short first, thus, last block will more likely contain 1 file only = faster depacking)
ReSort:
        Change = False
        For I As Integer = 0 To Prgs.Count - 2

            If FileLenA(I) > FileLenA(I + 1) Then
                PI = Prgs(I)
                Prgs(I) = Prgs(I + 1)
                Prgs(I + 1) = PI

                S = FileAddrA(I)
                FileAddrA(I) = FileAddrA(I + 1)
                FileAddrA(I + 1) = S

                S = FileOffsA(I)
                FileOffsA(I) = FileOffsA(I + 1)
                FileOffsA(I + 1) = S

                S = FileLenA(I)
                FileLenA(I) = FileLenA(I + 1)
                FileLenA(I + 1) = S

                IO = FileIOA(I)
                FileIOA(I) = FileIOA(I + 1)
                FileIOA(I + 1) = IO
                Change = True
            End If
        Next
        If Change = True Then GoTo ReSort

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddFileToPart()
        On Error GoTo Err

        Dim FN As String = ScriptEntryArray(0)
        Dim FA, FO, FL As String
        Dim FON, FLN As Integer
        Dim FUIO As Boolean = False

        Dim P() As Byte

        If Strings.Right(FN, 1) = "*" Then
            FN = Strings.Left(FN, Len(FN) - 1)
            FUIO = True
        End If

        'Get file variablea from script, or get default values if there were none in the script entry
        If IO.File.Exists(FN) = True Then
            P = IO.File.ReadAllBytes(FN)

            FA = If(ScriptEntryArray.Count > 1, ScriptEntryArray(1), ConvertNumberToHexString(P(0), P(1)))
            FO = If(ScriptEntryArray.Count > 2, ScriptEntryArray(2), "0002")
            FL = If(ScriptEntryArray.Count = 4, ScriptEntryArray(3), ConvertNumberToHexString((P.Length - 2) Mod 256, Int((P.Length - 2) / 256)))

            FON = ConvertHexStringToNumber(FO)
            FLN = ConvertHexStringToNumber(FL)

            'Make sure file length is not longer than actual file (should not happen)
            If FON + FLN > P.Length Then
                FLN = P.Length - FON
                FL = ConvertNumberToHexString(FLN Mod 256, Int(FLN / 256))
            End If

            'Trim file to the specified chunk (FLN number of bytes starting at FON, to Address of FAN)
            For I As Integer = 0 To FLN - 1
                P(I) = P(FON + I)
            Next
            ReDim Preserve P(FLN - 1)

        Else

            MsgBox("The following file does not exist:" + vbNewLine + vbNewLine + FN)
            Exit Sub

        End If

        FileCnt += 1
        ReDim Preserve FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt), FileIOA(FileCnt)

        FileAddrA(FileCnt) = FA
        FileOffsA(FileCnt) = FO     'This may not bee needed later
        FileLenA(FileCnt) = FL
        FileIOA(FileCnt) = FUIO

        If FirstFileOfDisk = True Then      'If Demo Start is not specified, we will use the start address of the first file
            FirstFileStart = FA
            FirstFileOfDisk = False
        End If

        Prgs.Add(P)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SplitScriptEntry()
        On Error GoTo Err

        If Strings.InStr(ScriptEntry, vbTab) = 0 Then
            ScriptEntryType = ScriptEntry
        Else
            ScriptEntryType = Strings.Left(ScriptEntry, Strings.InStr(ScriptEntry, vbTab) - 1)
            ScriptEntry = Strings.Right(ScriptEntry, ScriptEntry.Length - Strings.InStr(ScriptEntry, vbTab))
        End If

        LastNonEmpty = -1

        ReDim ScriptEntryArray(LastNonEmpty)

        ScriptEntryArray = Split(ScriptEntry, vbTab)

        For I As Integer = 0 To ScriptEntryArray.Length - 1
            If ScriptEntryArray(I) <> "" Then
                LastNonEmpty += 1
                ScriptEntryArray(LastNonEmpty) = ScriptEntryArray(I)
            End If
        Next

        If LastNonEmpty > -1 Then
            ReDim Preserve ScriptEntryArray(LastNonEmpty)
        Else
            ReDim ScriptEntryArray(0)
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ResetDiskVariables()
        On Error GoTo Err

        '-------------------------------------------------------------

        StartTrack = 1 : StartSector = 0

        NextTrack = StartTrack
        NextSector = StartSector

        BlockCnt = 0

        'SM1 = 0
        'SM2 = 0
        'LM = 0

        FirstFileOfDisk = True  'To save Start Address of first file on disk if Demo Start is not specified

        '-------------------------------------------------------------

        D64Name = ""
        DiskHeader = "demo disk " + Year(Now).ToString
        DiskID = "sprkl"
        DemoName = "demo"
        DemoStart = ""

        PartCnt = -1

        '-------------------------------------------------------------

        ResetPartVariables()    'Also adds first part

        NewPart = False

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ResetPartVariables()
        On Error GoTo Err

        FileCnt = -1
        ReDim FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt)
        Prgs.Clear()

        PartCnt += 1
        BlockCnt = 0

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

End Module
