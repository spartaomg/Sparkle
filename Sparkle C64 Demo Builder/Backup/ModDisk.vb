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
	Public MaxSector As Byte = 18, LastSector, BlockCnt, Prg(), IOBit As Byte
	Public Track(35), CT, CS, CP, BufferCt As Integer
	Public TestDisk As Boolean = False
	Public StartTrack As Byte = 1
	Public StartSector As Byte = 0

	Public ByteSt(), BitSt(), Buffer(255), BitPos, LastByte, AdLo, AdHi, MaxType As Byte
	Public Match, MaxLit, MaxBit, MatchSave(), MaxSave, OrigBC, PrgLen, Distant As Integer
	Public MatchOffset(), MatchCnt, RLECnt, MatchLen(), MaxOffset, MaxLen, LitCnt, Bits, BuffAdd, PrgAdd As Integer
	Public DistAd(), DistLen(), DistSave(), DistCnt, DistBase As Integer
	Public DtPos, CmPos, CmLast, DtLen, MatchStart, ByteCt, BitCt As Integer

	Public Compression As String

	Public Script As String
	Public ScriptHeader As String = "[Sparkle Loader Script]"
	Public ScriptName As String
	Public DiskNo As Integer
	Public D64Path As String

	'-----THESE WILL BE REMOVED-----
	Public D64Name As String
	Public DiskName As String = "OMG " + Year(Now).ToString
	'-------------------------------

	Public DiskHeader As String = "OMG " + Year(Now).ToString
	Public DiskID As String = "sprkl"
	Public DiskComp As String
	Public DemoName As String
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
	Public MusicA(), MusicInitA(), MusicPlayA() As String
	Public FilesA(), FileStartA(), FileSizeA() As String
	Public DiskNoA(), DFDiskNoA(), DFPartNoA() As Integer
	Public FilesInPartA() As Integer
	Public PDiskNoA(), PSizeA() As Integer
	Public FDiskNoA(), FPartNoA(), FSizeA() As Integer
	Public TotalParts As Integer = 0
	Public NewFile As String

	Public CompMode As Byte = 0

	Public DiskBlockSize() As Integer
	Public FileBlockSize() As Integer
	Public FBSDisk() As Integer
	Public MusicBlockSize As Integer = 0

	Public bBuildDisk As Boolean = False

	Dim SS, SE As Integer
	Dim EndOfSection As Boolean = False
	Dim NewPart As Boolean = False
	Dim EntryType As String = ""
	Dim Entry As String = ""
	Dim EntryArray() As String
	Dim LastNonEmpty As Integer = -1

	Public LC(), NM, FM, LM As Integer
	Public SM1 As Integer = 0
	Public SM2 As Integer = 0
	Public OverMaxLit As Boolean = False

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
		MsgBox(ErrorToString(), vbOKOnly, "SetLastSector Error")

	End Sub

	Public Function ConvertScriptToArray() As Boolean

		ConvertScriptToArray = False

		SS = 1 : SE = 1

		FindString()

		If Entry <> ScriptHeader Then
			MsgBox("Invalid Loader Script file!", vbExclamation + vbOKOnly, "Invalid file")
			Exit Function
		End If

		ResetArrays()
		AddNewDiskToArray()

FindNext:
		FindString()
		FindEntryType()

		If SE < Script.Length Then GoTo FindNext

		ConvertScriptToArray = True

	End Function

	Public Sub ResetArrays()

		DiskCnt = -1

		ReDim DiskNoA(DiskCnt)
		ReDim D64NameA(DiskCnt), DiskHeaderA(DiskCnt), DiskIDA(DiskCnt), DemoNameA(DiskCnt), DemoStartA(DiskCnt)
		ReDim MusicA(DiskCnt), MusicInitA(DiskCnt), MusicPlayA(DiskCnt)
		ReDim DiskBlockSize(DiskCnt)

		PartCnt = -1
		ReDim PDiskNoA(PartCnt), PSizeA(PartCnt)

		FileCnt = -1
		ReDim FilesA(FileCnt), FileStartA(FileCnt), FileSizeA(FileCnt), DFDiskNoA(FileCnt), DFPartNoA(FileCnt)
		ReDim FDiskNoA(FileCnt), FPartNoA(FileCnt), FSizeA(FileCnt)
		ReDim FileBlockSize(FileCnt), FBSDisk(FileCnt)

		TotalParts = 0

		MusicBlockSize = 0

	End Sub

	Public Sub AddNewDiskToArray()

		DiskCnt += 1
		ReDim Preserve DiskNoA(DiskCnt)
		ReDim Preserve D64NameA(DiskCnt), DiskHeaderA(DiskCnt), DiskIDA(DiskCnt), DemoNameA(DiskCnt), DemoStartA(DiskCnt)
		ReDim Preserve MusicA(DiskCnt), MusicInitA(DiskCnt), MusicPlayA(DiskCnt)
		ReDim Preserve DiskBlockSize(DiskCnt)

		DiskNoA(DiskCnt) = DiskCnt
		CurrentDisk = DiskCnt

	End Sub

	Public Sub AddNewPartToArray(DiskIndex As Integer)

		PartCnt += 1
		ReDim Preserve FilesInPartA(PartCnt), PDiskNoA(PartCnt), PSizeA(PartCnt)

		PDiskNoA(PartCnt) = DiskIndex
		CurrentPart = PartCnt
		'TotalParts += 1

	End Sub

	Public Sub AddNewFileToArray(DiskIndex As Integer)

		'If MsgBox("Add " + NewFile + " to Disk " + DiskIndex.ToString + ", Part " + CurrentPart.ToString + "?", vbOKCancel + vbCancel, "Add new file?") = vbCancel Then Exit Sub

		FileCnt += 1
		ReDim Preserve FilesA(FileCnt), DFDiskNoA(FileCnt), DFPartNoA(FileCnt), FileStartA(FileCnt), FileSizeA(FileCnt)
		ReDim Preserve FDiskNoA(FileCnt), FPartNoA(FileCnt), FSizeA(FileCnt)
		ReDim Preserve FileBlockSize(FileCnt), FBSDisk(FileCnt)

		DFDiskNoA(FileCnt) = DiskIndex
		DFPartNoA(FileCnt) = CurrentPart
		FilesInPartA(CurrentPart) += 1            'Number
		CurrentFile = FileCnt

	End Sub

	Private Sub FindString()
		On Error GoTo Err

		If Strings.Mid(Script, SS, 1) = Chr(13) Then    'Check if this is an empty line indicating a new section
			NewPart = True
			SS = SS + 2                                 'Skip EOL bytes
			SE = SS + 1
		Else
			NewPart = False
		End If
NextChar:
		If Strings.Mid(Script, SE, 1) <> Chr(13) Then   'Look for EOL
			SE += 1                                     'Not EOL
			If SE <= Script.Length Then                 'Go to next char if we haven't reached the end of the script
				GoTo NextChar
			Else
				Entry = Strings.Mid(Script, SS, SE - SS) 'Reached end of script, finish this entry
			End If
		Else                                            'Found EOL
			Entry = Strings.Mid(Script, SS, SE - SS)    'Finish this entry
			SS = SE + 2                                 'Skip EOL bytes
			SE = SS + 1
		End If

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "FindString Error")

	End Sub

	Private Sub FindEntryType()

		If Strings.InStr(Entry, vbTab) = 0 Then
			EntryType = ""
		Else
			EntryType = Strings.Left(Entry, Strings.InStr(Entry, vbTab) - 1)
			Entry = Strings.Right(Entry, Entry.Length - Strings.InStr(Entry, vbTab))
		End If

		SplitEntry()

		Select Case EntryType
			Case "Path:"
				D64NameA(CurrentDisk) = EntryArray(0)
			Case "Header:"
				DiskHeaderA(CurrentDisk) = EntryArray(0)
			Case "ID:"
				DiskIDA(CurrentDisk) = EntryArray(0)
			Case "Name:"
				DemoNameA(CurrentDisk) = EntryArray(0)
			Case "Comp:"
				DiskComp = EntryArray(0)
				If DiskComp = "LZ4+RLE" Then
					CompMode = 1
				Else
					CompMode = 0
				End If
			Case "Start:"
				DemoStartA(CurrentDisk) = EntryArray(0)
			Case "Music:"
				If EntryArray(0) <> "" Then
					MusicA(CurrentDisk) = EntryArray(0)
					If EntryArray.Count > 1 Then MusicInitA(CurrentDisk) = EntryArray(1)  'If we do not have this info then we assume Init=start address
					If EntryArray.Count = 3 Then MusicPlayA(CurrentDisk) = EntryArray(2)  'If we do not have this info then we assume Play=start address+3
					TotalParts += 1
				End If
			Case "File:"
				AddFileToArray()
				FilesA(FileCnt) = EntryArray(0) 'New File's Path
				If EntryArray.Count > 1 Then FileStartA(FileCnt) = EntryArray(1)
				If EntryArray.Count = 3 Then FileSizeA(FileCnt) = EntryArray(2)
			Case "New Disk"
				AddNewDiskToArray()
			Case Else
		End Select

	End Sub

	Private Sub AddFileToArray()

		If NewPart = True Then
			AddNewPartToArray(CurrentDisk)
			TotalParts += 1
		End If

		AddNewFileToArray(CurrentDisk)

	End Sub

	Private Sub SplitEntry()

		LastNonEmpty = -1

		ReDim EntryArray(LastNonEmpty)

		EntryArray = Split(Entry, vbTab)

		For I As Integer = 0 To EntryArray.Length - 1
			If EntryArray(I) <> "" Then
				LastNonEmpty += 1
				EntryArray(LastNonEmpty) = EntryArray(I)
			End If
		Next

		If LastNonEmpty > -1 Then
			ReDim Preserve EntryArray(LastNonEmpty)
		Else
			ReDim EntryArray(0)
		End If

	End Sub

	Public Sub BuildDemoDisk(DiskIndex As Integer)
		Dim FirstTrack, FirstSector As Byte
		Dim TotalBC As Integer
		Dim FC, PC As Integer

		StartTrack = 1 : StartSector = 0

		NextTrack = StartTrack
		NextSector = StartSector

		BlockCnt = 0

		ReDim ByteSt(0)
		ResetBuffer()

		TotalBC = 0
		For PC = 0 To PartCnt
			For FC = 0 To FilesInPartA(PC)
				If FC = 0 Then
					ResetBuffer()
				Else
				End If
			Next
		Next

	End Sub

	Public Sub BuildDisk(DiskIndex As Integer)
		Dim FirstTrack, FirstSector As Byte
		Dim TotalBC As Integer
		StartTrack = 1 : StartSector = 0

		NextTrack = StartTrack
		NextSector = StartSector

		BlockCnt = 0

		'ReDim ByteSt(0)
		'ResetBuffer()

		SM1 = 0
		SM2 = 0

		If MusicA(DiskIndex) <> "" Then
			'CompressFile(MusicA(DiskIndex))
			'CloseBuff()
			'ByteSt(255) = BlockCnt
			'AddCompressedPart(StartTrack, StartSector)
			AddCompressedPrg(MusicA(DiskIndex), StartTrack, StartSector)
			Disk(Track(StartTrack) + StartSector * 256 + 255) = BufferCt
		End If

		Dim FC, PC As Integer

		'ReDim ByteSt(0)
		'ResetBuffer()

		'If FileCnt > -1 Then
		'For FC = 0 To FileCnt
		'If DFDiskNoA(FC) = DiskIndex Then
		'If DFPartNoA(FC) > PartsA.Count Then
		'ReDim Preserve PartsA(DFPartNoA(FC))
		'End If
		'PartsA(DFPartNoA(FC)) += 1
		'End If
		'Next
		'End If
		'TotalBC = 0
		'For PC = 0 To PartsA.Count
		'For FC = 0 To PartsA(PC)
		'If FC = 0 Then
		'ResetBuffer()
		'Else
		'End If
		'Next
		'Next


		PC = -1
		TotalBC = 0
		If FileCnt > -1 Then
			For FC = 0 To FileCnt
				If DFDiskNoA(FC) = DiskIndex Then   'Is this file on the current disk?
					If DFPartNoA(FC) <> PC Then     'Is this te first file of a new part?
						'Yes, first file of new part
						If PC > -1 Then
							'Close last buffer
							'CloseBuff()
							'Finish previous part by updating block count
							'ByteSt(255) = BlockCnt
							Disk(Track(FirstTrack) + 256 * FirstSector + 255) = BlockCnt    'Write block number of previous part
							TotalBC += BlockCnt
						End If
						'Start next Part
						PC += 1
						FirstTrack = NextTrack
						FirstSector = NextSector
						BlockCnt = 0
					End If
					'Add file to part
					'If BlockCnt <> 0 Then AddDivider
					'CompressFile(FilesA(FC))
					AddCompressedPrg(FilesA(FC), NextTrack, NextSector)
				End If
			Next
			'CloseBuff()
			'ByteSt(255) = BlockCnt
			Disk(Track(FirstTrack) + 256 * FirstSector + 255) = BlockCnt
			'AddCompressedPart(NextTrack, NextSector)
			'MsgBox("BlockCnt:" + vbTab + BlockCnt.ToString)
			TotalBC += BlockCnt
		End If

		MsgBox("Block Count: " + TotalBC.ToString + " blocks" + vbNewLine +
			   "SM1:" + vbTab + SM1.ToString + vbNewLine +
			   "SM2:" + vbTab + SM2.ToString)

		SystemFile = True
		InjectDriveCode(0, TotalParts, 0)
		InjectLoader(DiskIndex, 18, 5, 6)
		SystemFile = False

		UpdateBAM(DiskIndex)

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "BuildDisk Error")

	End Sub

	Public Sub CompressFile(PN As String)

		ReDim Prg(0)

		If Strings.Right(PN, 1) = "*" Then
			Prg = IO.File.ReadAllBytes(Left(PN, Len(PN) - 1))
			IOBit = 1
		Else
			Prg = IO.File.ReadAllBytes(PN)
			IOBit = 0
		End If

		'If CompMode = 0 Then
		LZR2()

	End Sub

	Public Sub AddCompressedPart(Optional T As Byte = 1, Optional S As Byte = 0, Optional IL As Byte = 4, Optional AddToDir As Boolean = False, Optional PrgName As String = "", Optional UpdateBlocksFree As Boolean = True)

		If BlocksFree < BlockCnt Then
			MsgBox("Unable to add part to disk", vbOKOnly, "Not enough free space on disk")
			Exit Sub
		End If

		CT = T
		CS = S
		'MsgBox(CS.ToString + ":" + CT.ToString)
		If SectorOK(CT, CS) = False Then
			FindNextFreeSector()
		End If
		'MsgBox(BufferCt.ToString + vbNewLine + ByteSt.Length.ToString)
		For I = 0 To BufferCt - 1
			For J = 0 To 255
				'Disk(Track(CT) + 256 * CS + J) = Prg(I * 256 + J)
				Disk(Track(CT) + 256 * CS + J) = ByteSt(I * 256 + J)
			Next

			DeleteBit(CT, CS, UpdateBlocksFree)

			If AddInterleave() = False Then   'No free sectors left on disk
				If I = BufferCt - 1 Then
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

		'MsgBox(PN + vbNewLine + BufferCt.ToString + vbNewLine + BlockCnt.ToString)

		If AddToDir = True Then
			AddFileToDir(PrgName)
		End If

		ReDim ByteSt(0)

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "AddCompressedPart Error")

	End Sub

	Public Sub AddCompressedPrg(PN As String, Optional T As Byte = 1, Optional S As Byte = 0, Optional IL As Byte = 4, Optional AddToDir As Boolean = False, Optional PrgName As String = "", Optional UpdateBlocksFree As Boolean = True)
		'On Error GoTo Err
		'Dim BehindIO As Boolean

		ReDim Prg(0)

		If Strings.Right(PN, 1) = "*" Then
			Prg = IO.File.ReadAllBytes(Left(PN, Len(PN) - 1))
			IOBit = 1
		Else
			Prg = IO.File.ReadAllBytes(PN)
			IOBit = 0
		End If

		'If CompMode = 0 Then
		LZR2()
		'Else
		'LZ4RLE()
		'End If

		If BlocksFree < BufferCt Then
			MsgBox("Unable to add " + PN + " to disk", vbOKOnly, "Not enough free space on disk")
			Exit Sub
		End If

		CT = T
		CS = S
		'MsgBox(CS.ToString + ":" + CT.ToString)
		If SectorOK(CT, CS) = False Then
			FindNextFreeSector()
		End If
		'MsgBox(BufferCt.ToString + vbNewLine + ByteSt.Length.ToString)
		For I = 0 To BufferCt - 1
			For J = 0 To 255
				'Disk(Track(CT) + 256 * CS + J) = Prg(I * 256 + J)
				Disk(Track(CT) + 256 * CS + J) = ByteSt(I * 256 + J)
			Next

			DeleteBit(CT, CS, UpdateBlocksFree)

			If AddInterleave() = False Then   'No free sectors left on disk
				If I = BufferCt - 1 Then
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

		'MsgBox(PN + vbNewLine + BufferCt.ToString + vbNewLine + BlockCnt.ToString)

		If AddToDir = True Then
			AddFileToDir(PrgName)
		End If

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "AddCompressedPrg Error")

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
		MsgBox(ErrorToString(), vbOKOnly, "SectorOK Error")

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
		MsgBox(ErrorToString(), vbOKOnly, "FindNextFreeSector Error")

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
		MsgBox(ErrorToString(), vbOKOnly, "DeleteBit Error")

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
		MsgBox(ErrorToString(), vbOKOnly, "AddInterleave Error")

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
		MsgBox(ErrorToString(), vbOKOnly, "TrackIsFull Error")

	End Function

	Private Sub CalcNextSector(Optional IL As Byte = 5)

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

	End Sub

	Private Sub AddFileToDir(PrgName As String)
		On Error GoTo Err

		Dim Cnt, B, W As Integer
		Dim ST, SS, A As Byte

		CT = 18 : CS = 1
		Cnt = Track(CT) + CS * 256
SeekLastDirBlock:
		If Disk(Cnt) <> 0 Then
			Cnt = Track(Disk(Cnt)) + Disk(Cnt + 1) * 256    'Find last Directory block
			GoTo SeekLastDirBlock
		Else
LastDirBlock:
			B = 2
SeekNewEntry:
			If Disk(Cnt + B) = &H0 Then
				Disk(Cnt + B) = &H82
				Disk(Cnt + B + 1) = ST
				Disk(Cnt + B + 2) = SS
				For W = 0 To 15
					Disk(Cnt + B + 3 + W) = &HA0
				Next
				For W = 0 To Strings.Len(PrgName) - 1
					A = Asc(Strings.Mid(PrgName, W + 1, 1))
					If A > &H5F Then A -= &H20
					Disk(Cnt + B + 3 + W) = A
				Next
				Disk(Cnt + B + &H1C) = Int((Prg.Length + 2) / 256) + 1
			Else
				B = B + 32
				If B < 256 Then
					GoTo SeekNewEntry
				Else
					If AddInterleave(5) = False Then
						MsgBox("Unable to add the following entry to directory:" + vbNewLine + vbNewLine + PrgName, vbOKOnly, "Directory is full")
						Exit Sub
					End If
					DeleteBit(CT, CS, False)
					'CS += 5
					'If CS > 18 Then CS -= 18
					Disk(Cnt) = CT            'Add new track number to full Dir block
					Disk(Cnt + 1) = CS        'Add new sector number to full Dir block
					Cnt = Track(CT) + CS * 256  'Go to new T:S
					Disk(Cnt) = 0             'Mark it as last T:S in chain
					Disk(Cnt + 1) = 255
					GoTo LastDirBlock
				End If
			End If
		End If

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "AddFileToDir Error")

	End Sub

	Public Sub InjectDriveCode(Optional idcDiskID As Byte = 0, Optional idcFileCnt As Byte = 0, Optional idcNextID As Byte = 0)
		On Error GoTo Err

		Dim I, Cnt As Integer

		Drive = My.Resources.S2D

		Drive(802) = TotalParts 'Save number of parts to be loaded
		Drive(803) = NextID     'Save Next Side ID1

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

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "InjectDrivecode Error")

	End Sub

	Public Sub InjectLoader(DiskIndex As Integer, T As Byte, S As Byte, IL As Byte)
		On Error GoTo Err

		Dim B, I, Cnt, W As Integer
		Dim ST, SS, A, L, AdLo, AdHi As Byte
		Dim Loader() As Byte

		B = ConvertHexStringToNumber(DemoStartA(DiskIndex))
		AdLo = B Mod 256
		AdHi = Int(B / 256)

		If CompMode = 0 Then
			Loader = My.Resources.S2L_LZ    'LZ4 compression
		Else
			Loader = My.Resources.S2L_LR    'LZ4+RLE compression
		End If

		For I = 0 To Loader.Length - 3       'Find JMP $EDDA instruction
			If (Loader(I) = &H4C) And (Loader(I + 1) = &HDA) And (Loader(I + 2) = &HED) Then
				Loader(I + 1) = AdLo       'Lo Byte return address at the end of Loader
				Loader(I + 2) = AdHi       'Hi Byte return address at the end of Loader
				Exit For
			End If
		Next

		If MusicInitA(DiskIndex) <> "" Then
			B = ConvertHexStringToNumber(MusicInitA(0))
			AdLo = B Mod 256
			AdHi = Int(B / 256)

			For I = 0 To Loader.Length - 3     'Find JSR $1000 instruction
				If (Loader(I) = &H20) And (Loader(I + 1) = &H0) And (Loader(I + 2) = &H10) Then
					Loader(I + 1) = AdLo       'Lo Byte for Music Init
					Loader(I + 2) = AdHi       'Hi Byte for Music Init
					Exit For
				End If
			Next
		End If

		If MusicPlayA(DiskIndex) <> "" Then
			B = ConvertHexStringToNumber(MusicPlayA(DiskIndex))
			AdLo = B Mod 256
			AdHi = Int(B / 256)

			For I = Loader.Length - 3 To 0 Step -1       'Find INC $d019 instruction
				If (Loader(I) = &HEE) And (Loader(I + 1) = &H19) And (Loader(I + 2) = &HD0) Then
					Loader(I - 2) = AdLo       'Lo Byte for Music Play
					Loader(I - 1) = AdHi       'Hi Byte for Music Play
					Exit For
				End If
			Next
		End If

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
				For W = 0 To Strings.Len(DemoNameA(DiskIndex)) - 1
					A = Asc(Strings.Mid(DemoNameA(DiskIndex), W + 1, 1))
					If A > &H5F Then A -= &H20
					Disk(Cnt + B + 3 + W) = A
				Next
				Disk(Cnt + B + &H1C) = L
			Else
				B = B + 32
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

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly, "InjectLoader Error")

	End Sub

	Public Function ConvertHexStringToNumber(HexString As String) As Integer
		On Error GoTo Err

		Dim B, C, I As Integer
		B = 0
		For I = 1 To 4
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
		MsgBox(ErrorToString(), vbOKOnly, "ConvertHexStringToNumber Error")

	End Function

	Public Function ConvertHexNumberToString(HLo As Byte, Optional HHi As Byte = 0) As String
		Dim S As String = ""
		Dim B As Byte

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

		ConvertHexNumberToString = S

	End Function

	Public Sub UpdateBAM(DiskIndex As Integer)

		Dim Cnt As Integer
		Dim B As Byte
		CP = Track(18)

		For Cnt = &H90 To &HAA  'Name, ID, DOS type
			Disk(CP + Cnt) = &HA0
		Next

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

	End Sub

	Public Function ConvertArrayToScript() As String

		Dim S As String
		Dim PC As Integer = -1
		S = ScriptHeader + vbNewLine + vbNewLine

		For i As Integer = 0 To DiskCnt
			'S += "Path:" + vbTab + D64NameA(i) + vbNewLine +
			'	"Comp:" + vbTab + DiskComp + vbNewLine +
			'	"Header:" + vbTab + DiskHeaderA(i) + vbNewLine +
			'	"Name:" + vbTab + DemoNameA(i) + vbNewLine +
			'	"Start:" + vbTab + DemoStartA(i) + vbNewLine + vbNewLine +
			'	"Music:" + vbTab + MusicA(i) + vbTab + MusicInitA(i) + vbTab + MusicPlayA(i)
			S += "Path:" + vbTab + D64NameA(i) + vbNewLine +
				"Comp:" + vbTab + DiskComp + vbNewLine +
				"Header:" + vbTab + DiskHeaderA(i) + vbNewLine
			If DiskIDA(i) <> "" Then
				S += "ID:" + vbTab + DiskIDA(i) + vbNewLine
			End If
			S += "Name:" + vbTab + DemoNameA(i) + vbNewLine +
				"Start:" + vbTab + DemoStartA(i) + vbNewLine + vbNewLine +
				"Music:" + vbTab + MusicA(i) + vbTab + MusicInitA(i) + vbTab + MusicPlayA(i)
			For j As Integer = 0 To FileCnt
				S += vbNewLine
				If DFPartNoA(j) <> PC Then
					S += vbNewLine
					PC = DFPartNoA(j)
				End If
				S += "File:" + vbTab + FilesA(j)
			Next
		Next

		ConvertArrayToScript = S

	End Function

End Module
