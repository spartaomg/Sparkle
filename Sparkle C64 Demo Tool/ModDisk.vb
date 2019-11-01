Friend Module ModDisk

    Public Drive() As Byte

    Public BlocksFree As Integer = 664

    Public Disk(174847), NextTrack, NextSector As Byte   'Next Empty Track and Sector
    Public MaxSector As Byte = 18, LastSector, Prg() As Byte

    Public BufferCnt As Integer = 0

    Public FileUnderIO As Boolean = False
    Public IOBit As Byte

    Public Track(35), CT, CS, CP, BufferCt, BlockCnt As Integer
    Public TestDisk As Boolean = False
    Public StartTrack As Byte = 1
    Public StartSector As Byte = 0

    Public ByteSt(), BitSt(), Buffer(255), BitPos, LastByte, AdLo, AdHi As Byte
    Public Match, MaxBit, MatchSave(), PrgLen, Distant As Integer
    Public MatchOffset(), MatchCnt, RLECnt, MatchLen(), MaxSave, MaxOffset, MaxLen, LitCnt, Bits, BuffAdd, PrgAdd As Integer
    Public DistAd(), DistLen(), DistSave(), DistCnt, DistBase As Integer
    Public DtPos, CmPos, CmLast, DtLen, MatchStart, BitCnt As Integer
    Public ByteCnt As Integer    'Points at the next empty byte in buffer
    Public MatchType(), MaxType As String

    Public LastByteCt As Integer = 255
    Public LastMOffset As Integer = 0
    Public LastMLen As Integer = 0
    Public LastPOffset As Integer = -1
    Public LastMType As String = ""

    Public PreMOffset As Integer = 0
    Public PostMOffset As Integer = 0

    Public BitsSaved As Integer = 0
    Public BytesSaved As Integer = 0

    Public PreOLMO As Integer = 0
    Public PostOLMO As Integer = 0

    Public Script As String
    Public ScriptHeader As String = "[Sparkle Loader Script]"
    Public ScriptName As String
    Public DiskNo As Integer

    '-----THESE WILL BE REMOVED-----
    Public D64Name As String '= My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '-------------------------------

    Public DiskHeader As String = "demo disk " + Year(Now).ToString
    Public DiskID As String = "sprkl"
    Public DemoName As String = "demo"
    Public DemoStart As String
    Public SystemFile As Boolean = False
    Public FileChanged As Boolean = False

    Public DiskCnt As Integer = -1
    Public PartCnt As Integer = -1
    Public FileCnt As Integer = -1
    Public CurrentDisk As Integer = -1
    Public CurrentPart As Integer = -1
    Public CurrentFile As Integer = -1

    Public D64NameA(), DiskHeaderA(), DiskIDA(), DemoNameA(), DemoStartA(), DirArtA() As String
    Public FileNameA(), FileAddrA(), FileOffsA(), FileLenA() As String
    Public Prgs As New List(Of Byte())
    Public FileIOA() As Boolean

    Public DiskNoA(), DFDiskNoA(), DFPartNoA(), DiskPartCntA(), DiskFileCntA() As Integer
    Public FilesInPartA() As Integer
    Public PDiskNoA(), PSizeA() As Integer
    Public FDiskNoA(), FPartNoA(), FSizeA() As Integer
    Public TotalParts As Integer = 0
    Public NewFile As String

    Public DiskSizeA() As Integer
    Public PartSizeA() As Integer
    Public FileSizeA() As Integer
    Public FBSDisk() As Integer
    Public PartByteCntA() As Integer
    Public PartBitCntA() As Integer
    Public PartBitPosA() As Byte
    Public UncomPartSize As Double = 0

    Public bBuildDisk As Boolean = False

    Public SS, SE As Integer
    Public NewPart As Boolean = False
    Public ScriptEntryType As String = ""
    Public ScriptEntry As String = ""
    Public ScriptEntryArray() As String
    Public LastNonEmpty As Integer = -1

    Public LC(), NM, FM, LM As Integer
    Public SM1 As Integer = 0
    Public SM2 As Integer = 0
    Public OverMaxLit As Boolean = False

    Dim FirstFileOfDisk As Boolean = False
    Dim FirstFileStart As String = ""

    Public TabT(663), TabS(663) As Byte
    Public BlockPtr As Integer = 255
    Public LastBlockCnt As Byte = 0
    Public LoaderParts As Integer = 1
    Public FilesInBuffer As Byte = 1

    Private DirTrack, DirSector, DirPos As Integer
    Public DirArt As String
    Private DirEntry As String = ""
    Private LastDirSector As Byte

    Public ScriptPath As String

    Public CmdLine As Boolean = False

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

    Public Sub ResetArrays()
        On Error GoTo Err

        DiskCnt = -1

        ReDim DiskNoA(DiskCnt), DiskPartCntA(DiskCnt), DiskFileCntA(DiskCnt)
        ReDim D64NameA(DiskCnt), DiskHeaderA(DiskCnt), DiskIDA(DiskCnt), DemoNameA(DiskCnt), DemoStartA(DiskCnt), DirArtA(DiskCnt)
        ReDim DiskSizeA(DiskCnt)

        PartCnt = -1
        ReDim PDiskNoA(PartCnt), PSizeA(PartCnt), PartSizeA(PartCnt), FilesInPartA(PartCnt)

        FileCnt = -1

        ReDim FileNameA(FileCnt), DFDiskNoA(FileCnt), DFPartNoA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt)
        ReDim FDiskNoA(FileCnt), FPartNoA(FileCnt), FSizeA(FileCnt)
        ReDim FileSizeA(FileCnt), FBSDisk(FileCnt)

        TotalParts = 0

        'MusicBlockSize = 0

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function FindNextScriptEntry() As Boolean
        On Error GoTo Err

        FindNextScriptEntry = True

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

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        FindNextScriptEntry = False

    End Function

    Public Sub SplitEntry()
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

        For Cnt = 4 To (36 * 4) - 1
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

        CP = Track(CT) + (256 * CS)
        Disk(CP + 1) = 255

        NextTrack = 1 : NextSector = 0

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

        BP = Track(18) + (T * 4) + 1 + Int(S / 8)
        BB = 255 - (2 ^ (S Mod 8))    '=0-7
        'MsgBox(Hex(BP) + " : " + Hex(BB))
        Disk(BP) = Disk(BP) And BB

        BP = Track(18) + (T * 4)
        Disk(BP) -= 1
        'MsgBox(Disk(BP).ToString)
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

    Public Function InjectDriveCode(idcDiskID As Byte, idcFileCnt As Byte, idcNextID As Byte, Optional TestDisk As Boolean = False) As Boolean
        On Error GoTo Err

        InjectDriveCode = True

        Dim I, Cnt As Integer

        If TestDisk = True Then
            Drive = My.Resources.SDT
        Else
            Drive = My.Resources.SD
        End If

        Drive(802) = idcFileCnt 'Save number of parts to be loaded to ZP Tab Location $20 (=$320+2), includes Address Bytes!!!
        Drive(803) = idcNextID  'Save Next Side ID1 to ZP Tab Location $21 (=$321+2), includes Address Bytes

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
        Disk(Track(18) + (0 * 256) + 255) = idcDiskID
        Disk(Track(18) + (0 * 256) + 254) = idcFileCnt
        Disk(Track(18) + (0 * 256) + 253) = idcNextID

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        InjectDriveCode = False

    End Function

    Public Function InjectLoader(DiskIndex As Integer, T As Byte, S As Byte, IL As Byte, Optional TestDisk As Boolean = False) As Boolean
        On Error GoTo Err

        InjectLoader = True

        Dim B, I, Cnt, W As Integer
        Dim ST, SS, A, L, AdLo, AdHi As Byte
        Dim Loader() As Byte
        Dim DN As String

        If DiskIndex > -1 Then
            B = Convert.ToInt32(DemoStartA(DiskIndex), 16)
        Else
            B = Convert.ToInt32(DemoStart, 16)
        End If

        If B = 0 Then B = Convert.ToInt32(FirstFileStart, 16)

        If B = 0 Then
            MsgBox("Unable to build demo disk." + vbNewLine + vbNewLine + "Missing start address", vbOKOnly)
            InjectLoader = False
            Exit Function
        End If

        AdLo = (B - 1) Mod 256
        AdHi = Int((B - 1) / 256)

        If TestDisk = False Then
            Loader = My.Resources.SL
        Else
            Loader = My.Resources.SLT
        End If

        For I = 0 To Loader.Length - 3       'Find JMP $01f0 instruction (JMP AltLoad)
            If (Loader(I) = &H4C) And (Loader(I + 1) = &H80) And (Loader(I + 2) = &H1) Then
                Loader(I - 2) = AdLo       'Lo Byte return address at the end of Loader
                Loader(I - 5) = AdHi       'Hi Byte return address at the end of Loader
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
                Disk(Cnt + B + &H1C) = L    'Length of boot loader in blocks
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

    Public Function ConvertNumberToHexString(HLo As Byte, Optional HHi As Byte = 0) As String
        On Error GoTo Err

        ConvertNumberToHexString = LCase(Hex(HLo + (HHi * 256)))

        If Len(ConvertNumberToHexString) < 4 Then
            ConvertNumberToHexString = Left("0000", 4 - Len(ConvertNumberToHexString)) + ConvertNumberToHexString
        End If

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

    Public Sub MakeTestDisk()
        On Error GoTo Err

        Dim B As Byte
        Dim TDiff As Integer = Track(19) - Track(18)
        Dim SMax As Integer

        DemoStart = "0820"
        DemoName = "sparkle test"
        DiskHeader = "spakle test"
        DiskID = " 2019"

        NewDisk()

        For T As Integer = 1 To 35
            Select Case T
                Case 1 To 17
                    SMax = 20
                Case 19 To 24
                    SMax = 18
                Case 25 To 30
                    SMax = 17
                Case 31 To 35
                    SMax = 16
            End Select
            If T <> 18 Then
                For S As Integer = 0 To SMax
                    For I As Integer = S To S + 255
                        B = I Mod 256
                        Disk(Track(T) + S * 256 + I - S) = B
                    Next
                    DeleteBit(T, S, True)
                Next
            End If
        Next

        InjectLoader(-1, 18, 5, 6, True)
        InjectDriveCode(1, 255, 1, True)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function BuildDemoFromScript(Optional SaveIt As Boolean = True) As Boolean
        On Error GoTo Err

        BuildDemoFromScript = True

        SS = 1 : SE = 1

        'Check if this is a valid script
        If FindNextScriptEntry() = False Then GoTo NoDisk

        If ScriptEntry <> ScriptHeader Then
            MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly)
            GoTo NoDisk
        End If

        DiskCnt = -1


NewDisk:
        'Reset Disk Variables
        If ResetDiskVariables() = False Then GoTo NoDisk

FindNext:
        If FindNextScriptEntry() = False Then GoTo NoDisk
        'Split String
        If SplitScriptEntry() = False Then GoTo NoDisk

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
            Case "DirArt:"
                If IO.File.Exists(ScriptEntryArray(0)) Then
                    DirArt = IO.File.ReadAllText(ScriptEntryArray(0))
                End If
            Case "File:"
                'Add files to part array, if new part, it will first sort files in last part then add previous part to disk
                If AddFile() = False Then GoTo NoDisk
            Case "New Disk"
                'If new disk, sort, compress and add last part, then update part count, add loader & drive code, save disk, then GoTo NewDisk
                If FinishDisk(False, SaveIt) = False Then GoTo NoDisk
                GoTo NewDisk
            Case Else
        End Select

        If SE < Script.Length Then GoTo FindNext

        'Last disk: sort, compress and add last part, then update part count, add loader & drive code, save disk, and we are done :)
        If FinishDisk(True, SaveIt) = False Then GoTo NoDisk

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        BuildDemoFromScript = False

    End Function

    Private Sub AddHeaderAndID()
        On Error GoTo Err

        Dim B As Byte

        CP = Track(18)

        For Cnt = &H90 To &HAA
            Disk(CP + Cnt) = &HA0
        Next

        If DiskHeader.Length > 16 Then
            DiskHeader = Left(DiskHeader, 16)
        End If

        For Cnt = 1 To Strings.Len(DiskHeader)
            B = Asc(Mid(DiskHeader, Cnt, 1))
            If B > &H5F Then B -= &H20
            Disk(CP + &H8F + Cnt) = B
        Next

        If DiskID.Length > 5 Then
            DiskID = Left(DiskID, 5)
        End If

        For Cnt = 1 To Len(DiskID)                  'Overwrites Disk ID and DOS type (5 characters max.)
            B = Asc(Mid(DiskID, Cnt, 1))
            If B > &H5F Then B -= &H20
            Disk(CP + &HA1 + Cnt) = B
        Next

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function FinishDisk(LastDisk As Boolean, Optional SaveIt As Boolean = True) As Boolean
        On Error GoTo Err

        FinishDisk = True
        If SortPart() = False Then GoTo NoDisk
        If CompressPart() = False Then GoTo NoDisk
        FinishPart(0, True)
        CloseBuff()
        'Now add compressed parts to disk
        If AddCompressedPartsToDisk() = False Then GoTo NoDisk
        AddHeaderAndID()
        If InjectLoader(-1, 18, 5, 6) = False Then GoTo NoDisk
        If InjectDriveCode(DiskCnt + 1, LoaderParts, IIf(LastDisk = False, DiskCnt + 2, 0)) = False Then GoTo NoDisk
        If DirArt <> "" Then AddDirArt()

        BytesSaved += Int(BitsSaved / 8)
        BitsSaved = BitsSaved Mod 8

        If SaveIt = True Then
            If SaveDisk() = False Then GoTo NoDisk
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        FinishDisk = False

    End Function

    Private Function SaveDisk() As Boolean
        On Error GoTo Err

        SaveDisk = True

        'If D64Name = "" Then D64Name = "Demo Disk" + IIf(DiskCnt > 0, " " + (CurrentDisk + 1).ToString, "") + ".d64"
        If D64Name = "" Then D64Name = "Demo Disk " + (DiskCnt + 1).ToString + ".d64"

        If InStr(D64Name, ":") = 0 Then
            D64Name = ScriptPath + D64Name
        End If


        If CmdLine = True Then
            'We are in command line, just save the disk
            IO.File.WriteAllBytes(D64Name, Disk)
        Else
            'We are in app mode, show dialog
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

        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        SaveDisk = False

    End Function

    Public Function CompressPart() As Boolean
        On Error GoTo Err

        CompressPart = True

        Dim PreBCnt As Integer = BufferCnt

        If Prgs.Count = 0 Then Exit Function        'GoTo NoComp DOES NOT WORK!!!

        'DO NOT RESET ByteSt AND BUFFER VARIABLES HERE!!!

        If (BufferCnt = 0) And (ByteCnt = 254) Then
        Else
            PrgAdd = Convert.ToInt32(FileAddrA(0), 16)
            PrgLen = Convert.ToInt32(FileLenA(0), 16)
            FinishPart(CheckIO(PrgLen - 1), False)
        End If

        NewPart = True

        For I As Integer = 0 To Prgs.Count - 1
            NewLZ(Prgs(I).ToArray, FileAddrA(I),, FileLenA(I), FileIOA(I))  'NewPart is TRUE FOR THE FIRST FILE
            If I < Prgs.Count - 1 Then FinishFile()
        Next

        LastBlockCnt = BlockCnt

        'IF THE WHOLE PART IS LESS THAN 1 BLOCK, THEN "IT DOES NOT COUNT", Part Counter WILL NOT BE INCREASED
        If PreBCnt = BufferCnt Then
            PartCnt -= 1
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoComp:
        CompressPart = False

    End Function

    Private Function AddFile() As Boolean
        On Error GoTo Err

        AddFile = True

        If NewPart = True Then
            'First finish last part, if it exists
            If PartCnt > 0 Then
                'Sort files in part
                If SortPart() = False Then GoTo NoDisk
                'Then compress files and add them to part
                If CompressPart() = False Then GoTo NoDisk     'THIS WILL RESET NewPart TO FALSE

            End If

            'Then reset part variables (file arrays, prg array, block cnt), increase part counter
            ResetPartVariables()

        End If

        'Then add file to part
        If AddFileToPart() = False Then GoTo NoDisk

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        AddFile = False

    End Function

    Public Function SortPart() As Boolean
        On Error GoTo Err

        SortPart = True

        If Prgs.Count = 1 Then Exit Function 'GoTo NoSort DOES NOT WORK!!!

        Dim Change As Boolean
        Dim FSO, FEO, FSI, FEI As Integer   'File Start and File End Outer loop/Inner loop
        Dim PO(), PI() As Byte
        Dim S As String
        Dim IO As Boolean

        '--------------------------------------------------------------------------------
        'Check files for overlap

        For O As Integer = 0 To Prgs.Count - 2
            FSO = Convert.ToInt32(FileAddrA(O), 16)              'Outer loop File Start
            FEO = FSO + Convert.ToInt32(FileLenA(O), 16) - 1     'Outer loop File End
            For I As Integer = O + 1 To Prgs.Count - 1
                FSI = Convert.ToInt32(FileAddrA(I), 16)          'Inner loop File Start
                FEI = FSI + Convert.ToInt32(FileLenA(I), 16) - 1 'Inner loop File End
                '----|------+---------|--------OR-------|------+---------|-----------------
                '    FSO    FSI       FEO               FSO    FEI       FEO
                If ((FSI >= FSO) And (FSI <= FEO)) Or ((FEI >= FSO) And (FEI <= FEO)) Then
                    Dim OLS As Integer = IIf(FSO >= FSI, FSO, FSI)  'Overlap Start address
                    Dim OLE As Integer = IIf(FEO <= FEI, FEO, FEI)  'Overlap End address

                    If (OLS >= &HD000) And (OLE <= &HDFFF) And (FileIOA(O) <> FileIOA(I)) Then
                        'Overlap is IO memory only and different IO status - NO OVERLAP
                    Else
                        If MsgBox("The following two files overlap in Part " + PartCnt.ToString + ":" _
                           + vbNewLine + vbNewLine + FileNameA(I) + " ($" + Hex(FSI) + " - $" + Hex(FEI) + ")" + vbNewLine + vbNewLine _
                           + FileNameA(O) + " ($" + Hex(FSO) + " - $" + Hex(FEO) + ")" + vbNewLine + vbNewLine + "Do you want to proceed?", vbYesNo + vbExclamation) = vbNo Then GoTo NoSort
                    End If
                End If
            Next
        Next

        '--------------------------------------------------------------------------------
        'Append adjacent files
Restart:
        Change = False

        For O As Integer = 0 To Prgs.Count - 2
            FSO = Convert.ToInt32(FileAddrA(O), 16)
            FEO = Convert.ToInt32(FileLenA(O), 16)
            For I As Integer = O + 1 To Prgs.Count - 1
                FSI = Convert.ToInt32(FileAddrA(I), 16)
                FEI = Convert.ToInt32(FileLenA(I), 16)

                If FSO + FEO = FSI Then
                    'Inner files follows outer file immediately
                    If (FSI <= &HD000) Or (FSI > &HDFFF) Then
                        'Append files as they meet outside IO memory
Append:                 PO = Prgs(O)
                        PI = Prgs(I)
                        ReDim Preserve PO(FEO + FEI - 1)

                        For J As Integer = 0 To FEI - 1
                            PO(FEO + J) = PI(J)
                        Next

                        Prgs(O) = PO
                        Change = True
                    Else
                        If FileIOA(O) = FileIOA(I) Then
                            'Files meet inside IO memory, append only if their IO status is the same
                            GoTo Append
                        End If
                    End If
                ElseIf FSI + FEI = FSO Then
                    'Outer file follows inner file immediately
                    If (FSO <= &HD000) Or (FSO > &HDFFF) Then
                        'Prepend files as they meet outside IO memory
Prepend:                PO = Prgs(O)
                        PI = Prgs(I)
                        ReDim Preserve PI(FEI + FEO - 1)

                        For J As Integer = 0 To FEO - 1
                            PI(FEI + J) = PO(J)
                        Next

                        Prgs(O) = PI

                        FileAddrA(O) = FileAddrA(I)

                        Change = True
                    Else
                        If FileIOA(O) = FileIOA(I) Then
                            'Files meet inside IO memory, prepend only if their IO status is the same
                            GoTo Prepend
                        End If
                    End If
                End If

                If Change = True Then
                    FEO += FEI
                    FileLenA(O) = ConvertNumberToHexString(FEO Mod 256, Int(FEO / 256))
                    For J As Integer = I To Prgs.Count - 2
                        FileNameA(J) = FileNameA(J + 1)
                        FileAddrA(J) = FileAddrA(J + 1)
                        FileOffsA(J) = FileOffsA(J + 1)     'this may not be needed later
                        FileLenA(J) = FileLenA(J + 1)
                        FileIOA(J) = FileIOA(J + 1)
                    Next
                    FileCnt -= 1
                    ReDim Preserve FileNameA(Prgs.Count - 2), FileAddrA(Prgs.Count - 2), FileOffsA(Prgs.Count - 2), FileLenA(Prgs.Count - 2)
                    ReDim Preserve FileIOA(Prgs.Count - 2)
                    Prgs.Remove(Prgs(I))
                    GoTo Restart
                End If
            Next
        Next

        '--------------------------------------------------------------------------------
        'Sort files by length (short files first, thus, last block will more likely contain 1 file only = faster depacking)
ReSort:
        Change = False
        For I As Integer = 0 To Prgs.Count - 2
            'Sort except if file length < 3, to allow for ZP relocation script hack
            If (Convert.ToInt32(FileLenA(I), 16) > Convert.ToInt32(FileLenA(I + 1), 16)) And (Convert.ToInt32(FileLenA(I), 16) > 2) Then
                PI = Prgs(I)
                Prgs(I) = Prgs(I + 1)
                Prgs(I + 1) = PI

                S = FileNameA(I)
                FileNameA(I) = FileNameA(I + 1)
                FileNameA(I + 1) = S

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

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoSort:
        SortPart = False

    End Function

    Public Function AddFileToPart() As Boolean
        On Error GoTo Err

        AddFileToPart = True

        Dim FN As String = ScriptEntryArray(0)
        Dim FA As String = ""
        Dim FO As String = ""
        Dim FL As String = ""
        Dim FAN As Integer = 0
        Dim FON As Integer = 0
        Dim FLN As Integer = 0
        Dim FUIO As Boolean = False

        Dim P() As Byte

        If Strings.Right(FN, 1) = "*" Then
            FN = Replace(FN, "*", "")
            'FN = Strings.Left(FN, Len(FN) - 1)
            FUIO = True
        End If

        If Strings.InStr(FN, ":") = 0 Then  'relative file path
            FN = ScriptPath + FN            'look for file in script's folder
        End If

        'Get file variables from script, or get default values if there were none in the script entry
        If IO.File.Exists(FN) = True Then
            P = IO.File.ReadAllBytes(FN)

            Select Case ScriptEntryArray.Count
                Case 1  'No parameters in script
                    If Strings.InStr(Strings.LCase(FN), ".sid") <> 0 Then   'SID file - read parameters from file
                        FA = ConvertNumberToHexString(P(P(7)), (P(P(7) + 1)))
                        FO = ConvertNumberToHexString(P(7) + 2)
                        FL = ConvertNumberToHexString((P.Length - P(7) - 2) Mod 256, Int((P.Length - P(7) - 2) / 256))
                    Else                                                    'Any other files
                        If P.Length > 2 Then                                'We have at least 3 bytes in the file
                            FA = ConvertNumberToHexString(P(0), P(1))       'First 2 bytes define load address
                            FO = "0002"                                     'Offset=2, Length=prg length-2
                            FL = ConvertNumberToHexString((P.Length - 2) Mod 256, Int((P.Length - 2) / 256))
                        Else                                                'Short file without paramters -> STOP
                            MsgBox("File parameters are needed for the following file:" + vbNewLine + vbNewLine + FN, vbCritical + vbOKOnly, "Missing file parameters")
                            GoTo NoDisk
                        End If
                    End If
                Case 2  'One parameter in script
                    FA = ScriptEntryArray(1)                                'Load address from script
                    FO = "0000"                                             'Offset will be 0, length=prg length
                    FL = ConvertNumberToHexString(P.Length Mod 256, Int(P.Length / 256))
                Case 3  'Two parameters in script
                    FA = ScriptEntryArray(1)                                'Load address from script
                    FO = ScriptEntryArray(2)                                'Offset from script
                    FON = Convert.ToInt32(FO, 16)                           'Make sure offset is valid
                    If FON > P.Length - 1 Then
                        FON = P.Length - 1                                  'If offset>prg length-1 then correct it
                        FO = ConvertNumberToHexString(FON Mod 256, Int(FON / 256))
                    End If                                                  'Length=prg length- offset
                    FL = ConvertNumberToHexString((P.Length - FON) Mod 256, Int((P.Length - FON) / 256))
                Case 4  'Three parameters in script
                    FA = ScriptEntryArray(1)
                    FO = ScriptEntryArray(2)
                    FON = Convert.ToInt32(FO, 16)                           'Make sure offset is valid
                    If FON > P.Length - 1 Then
                        FON = P.Length - 1                                  'If offset>prg length-1 then correct it
                        FO = ConvertNumberToHexString(FON Mod 256, Int(FON / 256))
                    End If                                                  'Length=prg length- offset
                    FL = ScriptEntryArray(3)
            End Select

            'FA = If(ScriptEntryArray.Count > 1, ScriptEntryArray(1), ConvertNumberToHexString(P(0), P(1)))
            'FO = If(ScriptEntryArray.Count > 2, ScriptEntryArray(2), "0002")
            'FL = If(ScriptEntryArray.Count = 4, ScriptEntryArray(3), ConvertNumberToHexString((P.Length - 2) Mod 256, Int((P.Length - 2) / 256)))
            'MsgBox(FA + vbNewLine + FO + vbNewLine + FL)

            FAN = Convert.ToInt32(FA, 16)
            FON = Convert.ToInt32(FO, 16)
            FLN = Convert.ToInt32(FL, 16)

            'Make sure file length is not longer than actual file (should not happen)
            If FON + FLN > P.Length Then
                FLN = P.Length - FON
                FL = ConvertNumberToHexString(FLN Mod 256, Int(FLN / 256))
            End If

            'Make sure file address+length<&H10000
            If FAN + FLN > &H10000 Then
                FLN = &H10000 - FAN
                FL = ConvertNumberToHexString(FLN Mod 256, Int(FLN / 256))
            End If

            'Trim file to the specified chunk (FLN number of bytes starting at FON, to Address of FAN)
            For I As Integer = 0 To FLN - 1
                P(I) = P(FON + I)
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

        UncomPartSize += Int(FLN / 256)
        If FLN Mod 256 <> 0 Then
            UncomPartSize += 1
        End If

        If FirstFileOfDisk = True Then      'If Demo Start is not specified, we will use the start address of the first file
            FirstFileStart = FA
            FirstFileOfDisk = False
        End If

        Prgs.Add(P)

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        AddFileToPart = False

    End Function

    Private Function SplitScriptEntry() As Boolean
        On Error GoTo Err

        SplitScriptEntry = True

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

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        SplitScriptEntry = False

    End Function

    Public Function ResetDiskVariables() As Boolean
        On Error GoTo Err

        ResetDiskVariables = True

        DiskCnt += 1
        ReDim Preserve DiskSizeA(DiskCnt)

        BufferCnt = 0

        ReDim ByteSt(-1)
        ResetBuffer()

        'Reset Disk image
        NewDisk()
        BlockPtr = 255
        '-------------------------------------------------------------

        StartTrack = 1 : StartSector = 0

        NextTrack = StartTrack
        NextSector = StartSector

        'BlockCnt = 0   'DONE IN ResetPartVariables

        'SM1 = 0
        'SM2 = 0
        'LM = 0

        BitsSaved = 0 : BytesSaved = 0

        FirstFileOfDisk = True  'To save Start Address of first file on disk if Demo Start is not specified

        '-------------------------------------------------------------

        D64Name = ""
        DiskHeader = "demo disk " + Year(Now).ToString
        DiskID = "sprkl"
        DemoName = "demo"
        DemoStart = ""
        DirArt = ""

        PartCnt = -1        'WILL BE INCREASED TO 0 IN ResetPartVariables
        LoaderParts = 1
        FilesInBuffer = 1
        '-------------------------------------------------------------

        If ResetPartVariables() = False Then GoTo NoDisk    'Also adds first part

        NewPart = False

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        ResetDiskVariables = False

    End Function

    Public Function ResetPartVariables() As Boolean
        On Error GoTo Err

        ResetPartVariables = True

        FileCnt = -1
        ReDim FileNameA(FileCnt), FileAddrA(FileCnt), FileOffsA(FileCnt), FileLenA(FileCnt), FileIOA(FileCnt)

        Prgs.Clear()

        PartCnt += 1

        TotalParts += 1
        ReDim Preserve PartSizeA(TotalParts)
        BlockCnt = 0

        UncomPartSize = 0

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        ResetPartVariables = False

    End Function

    Public Function AddCompressedPartsToDisk() As Boolean
        On Error GoTo Err

        AddCompressedPartsToDisk = True

        If BlocksFree < BufferCnt Then
            MsgBox("Unable to add part to disk", vbOKOnly, "Not enough free space on disk")
            GoTo NoDisk
        End If

        'CT = T
        'CS = S

        'If SectorOK(CT, CS) = False Then
        'FindNextFreeSector()
        'End If

        For I = 0 To BufferCnt - 1
            CT = TabT(I)
            CS = TabS(I)
            For J = 0 To 255
                Disk(Track(CT) + 256 * CS + J) = ByteSt(I * 256 + J)
            Next

            DeleteBit(CT, CS, True)
        Next

        If BufferCnt < 664 Then
            NextTrack = TabT(BufferCnt)     'CT
            NextSector = TabS(BufferCnt)    'CS
        Else
            NextTrack = 18
            NextSector = 0
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoDisk:
        AddCompressedPartsToDisk = False

    End Function

    Public Function AddDirArt() As Boolean
        On Error GoTo Err

        AddDirArt = True

        DirTrack = 18
        DirSector = 1
        DirPos = 0
        FindNextDirPos()
        DirEntry = ""
        For I As Integer = 1 To DirArt.Length
            DirEntry += Mid(DirArt, I, 1)
            If Mid(DirArt, I, 1) = Chr(10) Then
                If DirPos <> 0 Then AddDirEntry()
                FindNextDirPos()
            End If
        Next

        If (DirEntry <> "") And (DirPos <> 0) Then AddDirEntry()

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

NoDisk:

        AddDirArt = False

    End Function

    Private Sub FindNextDirPos()
        On Error GoTo Err

        DirPos = 0

FindNextEntry:

        For I As Integer = 2 To 255 Step 32
            If Disk(Track(DirTrack) + (DirSector * 256) + I) = 0 Then
                DirPos = I
                Exit Sub
            End If
        Next

        FindNextDirSector()

        If DirSector <> 0 Then GoTo FindNextEntry

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub FindNextDirSector()
        On Error GoTo Err

        'Sector order: 1,7,13,3,9,15,4,8,12,16

        LastDirSector = DirSector

        Select Case DirSector
            Case 1
                DirSector = 7
            Case 7
                DirSector = 13
            Case 13
                DirSector = 3
            Case 3
                DirSector = 9
            Case 9
                DirSector = 15
            Case 15
                DirSector = 4
            Case 4
                DirSector = 8
            Case 8
                DirSector = 12
            Case 12
                DirSector = 16
            Case 16
                DirSector = 0
        End Select

        Disk(Track(DirTrack) + (LastDirSector * 256)) = DirTrack
        Disk(Track(DirTrack) + (LastDirSector * 256) + 1) = DirSector

        If DirSector <> 0 Then
            Disk(Track(DirTrack) + (DirSector * 256)) = 0
            Disk(Track(DirTrack) + (DirSector * 256) + 1) = 255
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddDirEntry()
        On Error GoTo Err

        Disk(Track(DirTrack) + (DirSector * 256) + DirPos + 0) = &H82   '"PRG" -  all dir entries will point to first file in dir
        Disk(Track(DirTrack) + (DirSector * 256) + DirPos + 1) = 18     'Track 18 (track pointer of boot loader)
        Disk(Track(DirTrack) + (DirSector * 256) + DirPos + 2) = 5      'Sector 5 (sector pointer of boot loader)

        If Right(DirEntry, 1) = Chr(10) Then
            DirEntry = Left(DirEntry, DirEntry.Length - 1)
        End If

        If Right(DirEntry, 1) = Chr(13) Then
            DirEntry = Left(DirEntry, DirEntry.Length - 1)
        End If

        If DirEntry.Length > 16 Then
            DirEntry = Left(DirEntry, 16)
        ElseIf DirEntry.Length < 16 Then
            For I As Integer = DirEntry.Length + 1 To 16
                DirEntry += Chr(160)
            Next
        End If

        For I As Integer = 1 To 16
            Disk(Track(DirTrack) + (DirSector * 256) + DirPos + 2 + I) = Asc(Mid(UCase(DirEntry), I, 1))
        Next

        DirEntry = ""

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub SetScriptPath(Path As String)
        On Error GoTo Err

        ScriptName = Path
        ScriptPath = ScriptName

        If Path = "" Then Exit Sub

        For I As Integer = Path.Length - 1 To 0 Step -1
            If Right(ScriptPath, 1) <> "\" Then
                ScriptPath = Left(ScriptPath, ScriptPath.Length - 1)
            Else
                Exit For
            End If
        Next

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub SimplifyScript()
        On Error GoTo Err

        If InStr(Script, ScriptPath) <> 0 Then
            Dim S As String = "This script could be simplified using the relative path of the script as a base folder." + vbNewLine
            S += vbNewLine + "Do you want to simplify the script before saving it?" + vbNewLine + vbNewLine
            S += "(Note: the simplified script will only work if it remains in its current folder!)"
            If MsgBox(S, vbYesNo + vbInformation, "Simplify script?") = vbYes Then
                Script = Replace(Script, ScriptPath, "")
            End If
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

End Module
