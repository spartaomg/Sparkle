Friend Module ModBetterPacker

	Private FirstBlockOfNextFile As Boolean = False 'If true, this is the first block of next file in same buffer, Lit Selector Bit NOT NEEEDED
	Private NextFileInBuffer As Boolean = False     'Indicates whether the next file is added to the same buffer

	Private BlockUnderIO As Integer = 0
	Private AdLoPos As Byte, AdHiPos As Byte

	'Match offset and length is 1 based
	Private ReadOnly MaxOffset As Integer = 255 + 1 'Offset will be decreased by 1 when saved
	Private ReadOnly ShortOffset As Integer = 63 + 1

	Private ReadOnly MaxLongLen As Byte = 254 + 1   'Cannot be 255, there is an INY in the decompression ASM code, and that would make YR=#$00
	Private ReadOnly MaxMidLen As Byte = 61 + 1     'Cannot be more than 61 because 62=LongMatchTag, 63=NextFileTage
	Private ReadOnly MaxShortLen As Byte = 3 + 1    '1-3, cannot be 0 because it is preserved for EndTag

	'Literal length is 0 based
	Private ReadOnly MaxLitLen As Integer = 1 + 4 + 8 + 32 - 1  '=44 - this seems to be optimal, 1+4+8+16 and 1+4+8+64 are worse...

	Private MatchBytes As Integer = 0
	Private MatchBits As Integer = 0
	Private LitBits As Integer = 0
	Private MLen As Integer = 0
	Private MOff As Integer = 0

	Private MaxBits As Integer = 2048

	Structure Sequence
		Public Len As Integer           'Length of the sequence in bytes (0 based)
		Public Off As Integer           'Offset of Match sequence in bytes (1 based), 0 if Literal Sequence
		Public Bit As Integer           'Total Bits in Buffer
	End Structure

	Private Seq() As Sequence           'Sequence array, to find the best sequence
	Private SL(), SO(), LL(), LO() As Integer
	Private SI As Integer               'Sequence array index
	Private StartPos As Integer

	Public Sub PackFile(PN As Byte(), Optional FA As String = "", Optional FUIO As Boolean = False)
		On Error GoTo Err

		'----------------------------------------------------------------------------------------------------------
		'PROCESS FILE
		'----------------------------------------------------------------------------------------------------------

		'Dim PS As String = ""
		'For I As Integer = 0 To 255
		'PS += Hex(PN(I)) + vbTab
		'Next
		'MsgBox(PS)

		Prg = PN
		FileUnderIO = FUIO
		PrgAdd = Convert.ToInt32(FA, 16)
		PrgLen = Prg.Length

		ReDim SL(PrgLen - 1), SO(PrgLen - 1), LL(PrgLen - 1), LO(PrgLen - 1)
		ReDim Seq(PrgLen)       'This is actually one element more in the array, to have starter element with 0 values

		'LitCnt = 0              'Reset LitCnt: we start with 1 literal (LitCnt is 0 based)

		With Seq(1)             'Initialize first element of sequence
			'.Len = 0            '1 Literal byte, Len is 0 based
			'.Off = 0            'Offset=0 -> literal sequence, Off is 1 based
			.Bit = 10           'LitLen bit + 8 bits + type (Lit vs Match) selector bit 
		End With
		'----------------------------------------------------------------------------------------------------------
		'CALCULATE BEST SEQUENCE
		'----------------------------------------------------------------------------------------------------------

		CalcBestSequence(PrgLen - 1, 1)

		'----------------------------------------------------------------------------------------------------------
		'DETECT BUFFER STATUS AND INITIALIZE COMPRESSION
		'----------------------------------------------------------------------------------------------------------

		If ((BufferCnt = 0) And (ByteCnt = 254)) Or (ByteCnt = 255) Then
			FirstBlockOfNextFile = False                            'First block in buffer, Lit Selector Bit is needed (will be compression bit)
			NextFileInBuffer = False                                'This is the first file that is being added to an empty buffer
		Else
			FirstBlockOfNextFile = True                             'First block of next file in same buffer, Lit Selector Bit NOT NEEEDED
			NextFileInBuffer = True                                 'Next file is being added to buffer that already has data
		End If

		If NewPart Then
			BlockPtr = ByteSt.Count + 255                           'If this is a new part, store Block Counter Pointer
			NewPart = False
		End If

		Buffer(ByteCnt) = (PrgAdd + PrgLen - 1) Mod 256             'Add Address Hi Byte
		AdLoPos = ByteCnt

		If CheckIO(PrgLen - 1) = 1 Then                             'Check if last byte of block is under IO or in ZP
			BlockUnderIO = 1                                        'Yes, set BUIO flag
			ByteCnt -= 1                                            'And skip 1 byte (=0) for IO Flag
		Else
			BlockUnderIO = 0
		End If

		Buffer(ByteCnt - 1) = Int((PrgAdd + PrgLen - 1) / 256)      'Add Address Lo Byte
		AdHiPos = ByteCnt - 1

		ByteCnt -= 2
		LastByte = ByteCnt          'The first byte of the ByteStream after (BlockCnt and IO Flag and) Address Bytes (251..253)

		'----------------------------------------------------------------------------------------------------------
		'COMPRESS FILE
		'----------------------------------------------------------------------------------------------------------

		Pack()

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub CalcBestSequence(SeqStart As Integer, SeqEnd As Integer)
		On Error GoTo Err

		Dim MaxO, MaxL As Integer
		Dim SeqLen, SeqOff As Integer
		Dim LeastBits As Integer

		'----------------------------------------------------------------------------------------------------------
		'CALCULATE MAX MATCH LENGTHS AND OFFSETS FOR EACH POSITION
		'----------------------------------------------------------------------------------------------------------

		'Pos = Max to Min>0 value
		For Pos As Integer = SeqStart To SeqEnd Step -1  'Pos cannot be 0, Prg(0) is always literal as it is always 1 byte left
			SO(Pos) = 0
			SL(Pos) = 0
			LO(Pos) = 0
			LL(Pos) = 0
			'Offset goes from 1 to max offset (cannot be 0)
			MaxO = IIf(Pos + MaxOffset < SeqStart, MaxOffset, SeqStart - Pos)
			'Match length goes from 1 to max length
			MaxL = IIf(Pos >= MaxLongLen, MaxLongLen, Pos)  'MaxL=254 or less
			For O As Integer = 1 To MaxO                                    'O=1 to 255 or less
				'Check if first byte matches at offset, if not go to next offset
				If Prg(Pos) = Prg(Pos + O) Then
					For L As Integer = 1 To MaxL                            'L=1 to 254 or less
						If L = MaxL Then
							GoTo Match
						ElseIf Prg(Pos - L) <> Prg(Pos + O - L) Then
							'L=MatchLength + 1 here
							If L >= 2 Then
Match:                          If O <= ShortOffset Then
									If (SL(Pos) < MaxShortLen) And (SL(Pos) < L) Then
										SL(Pos) = IIf(L > MaxShortLen, MaxShortLen, L)
										SO(Pos) = O       'Keep O 1-based
									End If
									If LL(Pos) < L Then
										LL(Pos) = L
										LO(Pos) = O
									End If
								Else
									If (LL(Pos) < L) And (L > 2) Then 'Skip short (2-byte) Mid Matches
										LL(Pos) = L
										LO(Pos) = O
									End If
								End If
							End If
							Exit For
						End If
					Next
					'If both short and long matches maxed out, we can leave the loop and go to the next Prg position
					If (LL(Pos) = IIf(Pos >= MaxLongLen, MaxLongLen, Pos)) And
						(SL(Pos) = IIf(Pos >= MaxShortLen, MaxShortLen, Pos)) Then
						Exit For
					End If
				End If
			Next
		Next

		'----------------------------------------------------------------------------------------------------------
		'FIND BEST SEQUENCE FOR EACH POSITION
		'----------------------------------------------------------------------------------------------------------

		For Pos As Integer = SeqEnd To SeqStart     'Start with second element, first has been initialized  above
			LeastBits = &HFFFFFF                    'Max block size=100 = $10000 bytes = $80000 bits, make default larger than this

			If LL(Pos) <> 0 Then
				SeqLen = LL(Pos)
			ElseIf SL(Pos) <> 0 Then
				SeqLen = SL(Pos)
			Else
				'Both LL(Pos) and SL(Pos) are 0, so this is a literal byte
				GoTo Literals
			End If

			'Check all possible lengths
			For L As Integer = SeqLen To 2 Step -1
				'For L As Integer = SeqLen To IIf(SeqLen - 2 > 2, SeqLen - 2, 2) Step -1
				'Get offset, use short match if possible
				SeqOff = IIf(L <= SL(Pos), SO(Pos), LO(Pos))

				''THIS DOES NOT SEEM TO MAKE ANY DIFFERENCE. RATHER, WE ARE SIMPLY EXLUDING ANY 2-BYTE MID MATCHES
				'If (L = 2) And (SeqOff > ShortOffset) Then
				'If LO(Pos - 2) = 0 And LO(Pos + 1) = 0 And SO(Pos - 2) = 0 And SO(Pos + 1) = 0 Then
				''Filter out short mid matches surrounded by literals
				'GoTo Literals
				'End If
				'End If

				'Calculate MatchBits
				CalcMatchBits(L, SeqOff)

				'See if total bit count is better than best version
				If Seq(Pos + 1 - L).Bit + MatchBits < LeastBits Then
					'If better, update best version
					LeastBits = Seq(Pos + 1 - L).Bit + MatchBits
					'and save it to sequence at Pos+1 (position is 1 based)
					With Seq(Pos + 1)
						.Len = L            'MatchLen is 1 based
						.Off = SeqOff       'Off is 1 based
						.Bit = LeastBits
					End With
				End If
			Next

Literals:
			'Continue previous Lit sequence or start new sequence
			LitCnt = If(Seq(Pos).Off = 0, Seq(Pos).Len, -1)

			'Calculate literal bits for a presumtive LitCnt+1 value
			CalcLitBits(LitCnt + 1)             'This updates LitBits
			LitBits += (LitCnt + 2) * 8         'Lit Bits + Lit Bytes
			'See if total bit count is less than best version
			If Seq(Pos - LitCnt - 1).Bit + LitBits < LeastBits Then  '=Seq(Pos - (LitCnt + 1)) simplified
				'If better, update best version
				LeastBits = Seq(Pos - LitCnt - 1).Bit + LitBits  '=Seq(Pos - (LitCnt + 1)) simplified
				'and save it to sequence at Pos+1 (position is 1 based)
				With Seq(Pos + 1)
					.Len = LitCnt + 1       'LitCnt is 0 based, LitLen is 0 based
					.Off = 0                'An offset of 0 marks a literal sequence, match offset is 1 based
					.Bit = LeastBits
				End With
			End If

		Next

		'Dim S As String = ""

		'If IO.File.Exists(UserDeskTop + "\Seq.txt") = False Then

		'For I As Integer = SeqEnd To SeqStart
		'S += I.ToString + vbTab + SL(I).ToString + vbTab + SO(I).ToString + vbTab + LL(I).ToString + vbTab + LO(I).ToString + vbTab +
		'						(I + 1).ToString + vbTab + Seq(I + 1).Len.ToString + vbTab + Seq(I + 1).Off.ToString + vbTab + Seq(I + 1).Bit.ToString + vbNewLine
		'Next

		'IO.File.WriteAllText(UserDeskTop + "\Seq.txt", S)
		'End If

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub Pack()
		On Error GoTo Err

		Dim BufferFull As Boolean

		SI = PrgLen - 1
		StartPos = SI

Restart:
		Do

			If Seq(SI + 1).Off = 0 Then
				'--------------------------------------------------------------------
				'Literal sequence
				'--------------------------------------------------------------------
				LitCnt = Seq(SI + 1).Len                'LitCnt is 0 based
				MLen = 0                                'Reset MLen - this is needed for accurate bit counting in SequenceFits

				BufferFull = False
				Do While LitCnt > -1
					If SequenceFits(LitCnt + 1, CalcLitBits(LitCnt), CheckIO(SI - LitCnt)) = True Then
						'IDENTIFY 2-BYTE MIDMATCHES THAT MAY SAVE A FEW BITS
						'Select Case LitCnt Mod (MaxLitLen + 1)
						'Case 1, 2, 5, 13
						'If LitCnt > 0 Then
						'If FindShortMidMatch() = True Then
						'Exit Do
						'End If
						'End If
						'End Select
						AddLitBytes(LitCnt)
						Exit Do
					End If
					If LitCnt > Int(MaxBits / 8) Then   'Bypass too large numbers to improve speed
						LitCnt = Int(MaxBits / 8)
					End If
					LitCnt -= 1
					BufferFull = True
				Loop

				'Go to next element in sequence
				SI -= LitCnt + 1    'If nothing added to the buffer, LitCnt=-1+1=0

				If BufferFull = True Then
					AddLitBits()    'Add literal bits of the last literal sequence
					CloseBuffer()   'The whole literal sequence did not fit, buffer is full, close it
				End If

			Else
				'--------------------------------------------------------------------
				'Match sequence
				'--------------------------------------------------------------------

				BufferFull = False

				MLen = Seq(SI + 1).Len      '1 based
				MOff = Seq(SI + 1).Off      '1 based
Match:
				CalcMatchBits(MLen, MOff)
				If MatchBytes = 3 Then
					'--------------------------------------------------------------------
					'Long Match - 3 match bytes
					'--------------------------------------------------------------------
					If SequenceFits(3, CalcLitBits(LitCnt), CheckIO(SI - MLen + 1)) Then
						AddLitBits()
						'Add long match
						AddLongMatch()
					Else
						MLen = MaxMidLen
						BufferFull = True   'Buffer if full, we will need to close it
						GoTo CheckMid
					End If
				ElseIf MatchBytes = 2 Then
					'--------------------------------------------------------------------
					'Mid Match - 2 match bytes
					'--------------------------------------------------------------------
CheckMid:           If SequenceFits(2, CalcLitBits(LitCnt), CheckIO(SI - MLen + 1)) Then
						AddLitBits()
						'Add mid match
						AddMidMatch()
					Else
						BufferFull = True
						If SO(SI) <> 0 Then
							MLen = SL(SI)   'SL and SO array indeces are 0 based
							MOff = SO(SI)
							GoTo CheckShort
						Else
							GoTo CheckLit
						End If  'Short vs Literal
					End If      'Mid vs Short
				Else
					'--------------------------------------------------------------------
					'Short Match - 1 match byte
					'--------------------------------------------------------------------
CheckShort:         If SequenceFits(1, CalcLitBits(LitCnt), CheckIO(SI - MLen + 1)) Then
						AddLitBits()
						'Add short match
						AddShortMatch()
					Else
						'--------------------------------------------------------------------
						'Match does not fit, check if 1 literal byte fits
						'--------------------------------------------------------------------
						BufferFull = True
CheckLit:               MLen = 0    'This is needed here for accurate Bit count calculation in SequenceFits (indicates Literal, not Match)
CheckNextLit:           If SequenceFits(1, CalcLitBits(LitCnt + 1), CheckIO(SI - LitCnt)) Then
							LitCnt += 1     '0 based
							AddLitBytes(0)   'Add 1 literal byte (the rest has been added previously)
							MLen += 1       '1 based
						End If  'Literal vs nothing
					End If      'Short match vs literal
				End If          'Long, mid, or short match
Done:
				SI -= MLen

				If BufferFull Then
					AddLitBits()
					CloseBuffer()
				End If
			End If              'Lit vs match

		Loop While SI >= 0

		AddLitBits()            'See if any literal bits need to be added, space has been previously reserved for them

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Function FindShortMidMatch() As Boolean
		On Error GoTo Err

		FindShortMidMatch = False

		Dim StopP As Integer = If(StartPos - SI > MaxOffset, MaxOffset, StartPos - SI) - 1

		For I As Integer = 1 To StopP
			If (Prg(SI) = Prg(SI + I)) And (Prg(SI - 1) = Prg(SI + I - 1)) Then

				MLen = 2
				MOff = I
				LitCnt = -1
				If MOff > ShortOffset Then
					AddMidMatch()
				Else
					AddShortMatch()     'This should not happen...
				End If
				SI -= 2
				FindShortMidMatch = True
				Exit For
			End If
		Next

		Exit Function
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Function CalcMatchBits(Length As Integer, Offset As Integer) As Integer 'Match Length is 1 based
		On Error GoTo Err

		If (Length <= MaxShortLen) And (Offset <= ShortOffset) Then
			MatchBytes = 1
			'MatchBits = 8 + 1       '1 match byte + 1 type selector bit AFTER match sequence
		ElseIf Length <= MaxMidLen Then
			MatchBytes = 2
			'MatchBits = 16 + 1      '2 match bytes + 1 type selector bit AFTER match sequence
		Else
			MatchBytes = 3
			'MatchBits = 24 + 1      '3 match bytes + 1 type selector bit AFTER match sequence
		End If

		MatchBits = (MatchBytes * 8) + 1

		CalcMatchBits = MatchBits

		Exit Function
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Function CalcLitBits(Lits As Integer) As Integer     'LitCnt is 0 based
		On Error GoTo Err

		If Lits = -1 Then
			CalcLitBits = 0 + 1                     '0	1 type selector bit for match
		Else
			CalcLitBits = Fix(Lits / (MaxLitLen + 1)) * 9

			Select Case Lits Mod (MaxLitLen + 1)
				Case 0
					CalcLitBits += 1 + 0           '1	1 bittab bit
				Case 1 To 4
					CalcLitBits += 2 + 2 + 0       '4	2 bittab bits + 2 lit sequence length bits
				Case 5 To 12
					CalcLitBits += 3 + 3 + 0       '6	3 bittab bits + 3 lit sequence length bits
				Case 13 To MaxLitLen - 1
					CalcLitBits += 3 + 5 + 0       '8	3 bittab bits + 5 lit sequence length bits
				Case MaxLitLen
					CalcLitBits += 3 + 5 + 1       '9	3 bittab bits + 5 lit sequence length bits + 1 type selector bit AFTER lit sequence
			End Select
		End If

		LitBits = CalcLitBits

		Exit Function
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Function

	Private Function SequenceFits(BytesToAdd As Integer, BitsToAdd As Integer, Optional SequenceUnderIO As Integer = 0) As Boolean
		On Error GoTo Err

		'Calculate total bit count in buffer from ByteCnt, BitCnt, and BitPos
		Dim BitsInBuffer As Integer = ((255 - ByteCnt) * 8) + (BitCnt * 8) + 16 - BitPos
		MaxBits = 2048 - BitsInBuffer

		'Add Close Byte + IO Byte ONLY if this is the first sequence in the block that goes under IO
		BytesToAdd += 1 + If((BlockUnderIO = 0) And (SequenceUnderIO = 1), 1, 0)

		'Add Close Bit if this sequence is a match
		BitsToAdd += If(MLen > 1, 1, 0)

		'Check if sequence will fit within block size limits
		If BitsInBuffer + (BytesToAdd * 8) + BitsToAdd <= 2048 Then
			SequenceFits = True
			'Data will fit
			If (BlockUnderIO = 0) And (SequenceUnderIO = 1) Then
				'This is the first byte in the block that will go UIO, so lets update the buffer to include the IO flag
				For I As Integer = ByteCnt To AdHiPos   'Move all data to the left in buffer, including AdHi
					Buffer(I - 1) = Buffer(I)
				Next
				Buffer(AdHiPos) = 0                     'IO Flag to previous AdHi Position
				ByteCnt -= 1                            'Update ByteCt to next empty position in buffer
				LastByteCt -= 1                         'Last Match pointer also needs to be updated (BUG FIX - REPORTED BY RAISTLIN/G*P)
				AdHiPos -= 1                            'Update AdHi Position in Buffer
				BlockUnderIO = 1                        'Set BlockUnderIO Flag
			End If
		Else
			SequenceFits = False
		End If

		Exit Function
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

		SequenceFits = False

	End Function

	Private Sub AddLongMatch()
		On Error GoTo Err

		TotMatch += 1

		If (LitCnt = -1) Or (LitCnt Mod (MaxLitLen + 1) = MaxLitLen) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		Buffer(ByteCnt) = LongMatchTag                   'Long Match Flag = &HF8
		Buffer(ByteCnt - 1) = MLen - 1
		Buffer(ByteCnt - 2) = MOff - 1
		ByteCnt -= 3

		LitCnt = -1

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddMidMatch()
		On Error GoTo Err

		TotMatch += 1

		If (LitCnt = -1) Or (LitCnt Mod (MaxLitLen + 1) = MaxLitLen) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		Buffer(ByteCnt) = (MLen - 1) * 4                         'Length of match (#$02-#$3f, cannot be #$00 (end byte), and #$01 - distant selector??)
		Buffer(ByteCnt - 1) = MOff - 1
		ByteCnt -= 2

		LitCnt = -1

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddShortMatch()
		On Error GoTo Err

		TotMatch += 1

		If (LitCnt = -1) Or (LitCnt Mod (MaxLitLen + 1) = MaxLitLen) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		Buffer(ByteCnt) = ((MOff - 1) * 4) + (MLen - 1)
		ByteCnt -= 1

		LitCnt = -1

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddLitBytes(Lits As Integer)
		On Error GoTo Err

		For I As Integer = 0 To Lits
			Buffer(ByteCnt) = Prg(SI - I)   'Add byte to Byte Stream
			ByteCnt -= 1                    'Update Byte Position Counter
		Next I

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddLitBits()
		On Error GoTo Err

		If LitCnt = -1 Then Exit Sub    'We only call this routine with LitCnt>-1

		TotLit += Int(LitCnt / (MaxLitLen + 1)) + 1

		For I As Integer = 1 To Fix(LitCnt / (MaxLitLen + 1))
			If FirstBlockOfNextFile = True Then
				AddRBits(&B11111111, 8)
				FirstBlockOfNextFile = False
			Else
				AddRBits(&B111111111, 9)
			End If
		Next

		If FirstBlockOfNextFile = False Then
			AddRBits(1, 1)               'Add Literal Selector if this is not the first (Literal) byte in the buffer
		Else
			FirstBlockOfNextFile = False
		End If

		Dim Lits As Integer = LitCnt Mod (MaxLitLen + 1)

		Select Case Lits
			Case 0
				AddRBits(0, 1)              'Add Literal Length Selector 0	- read no more bits
			Case 1 To 4
				AddRBits(2, 2)              'Add Literal Length Selector 10 - read 2 more bits
				AddRBits(Lits - 1, 2)       'Add Literal Length: 00-03, 2 bits	-> 1000 00xx when read
			Case 5 To 12
				AddRBits(6, 3)              'Add Literal Length Selector 110 - read 3 more bits
				AddRBits(Lits - 5, 3)       'Add Literal Length: 00-07, 3 bits	-> 1000 1xxx when read
			Case 13 To MaxLitLen - 1
				AddRBits(7, 3)              'Add Literal Length Selector 111 - read 5 more bits
				AddRBits(Lits - 13, 5)      'Add Literal Length: 00-1f, 5 bits	-> 101x xxxx when read
			Case MaxLitLen
				AddRBits(7, 3)              'Add Literal Length Selector 111 - read 5 more bits
				AddRBits(Lits - 13, 5)      'Add Literal Length: 00-1f, 5 bits	-> 101x xxxx when read
				'AddRBits(0, 1)              'Add Match Selector Bit - not here, this is done at adding matches
		End Select

		'DO NOT RESET LitCnt HERE!!!

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Private Sub AddRBits(Bit As Integer, BCnt As Byte)
		On Error GoTo Err

		For I As Integer = BCnt - 1 To 0 Step -1
			If (Bit And 2 ^ I) <> 0 Then
				Buffer(BitCnt) = Buffer(BitCnt) Or 2 ^ (BitPos - 8)
			End If
			BitPos -= 1
			If BitPos < 8 Then
				BitPos += 8
				BitCnt += 1
			End If
		Next

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Public Sub CloseBuffer()
		On Error GoTo Err

		Buffer(ByteCnt) = EndTag            'Technically, not needed, the default value is #$00 anyway
		Buffer(0) = Buffer(0) And &H7F      'Delete Compression Bit (Default (i.e. compressed) value is 0)

		'FIND UNCOMPRESSIBLE BLOCKS (only used if there is ONE file in the BLOCK)
		'THE COMPRESSION BITs ADD 80.5 BYTES TO THE DISK, BUT THIS MAKES DEPACKING OF UNCOMPRESSIBLE BLOCKS MUCH FASTER
		If (StartPos - SI <= LastByte) And (StartPos > LastByte - 1) And (NextFileInBuffer = False) Then
			'if (StartPos - SI <= 100) And (StartPos > LastByte - 1) And (NextFileInBuffer = False) Then

			'Less than 252/253 bytes   AND  Not the end of File      AND  No other files in this buffer
			LastByte = AdLoPos - 2

			'Check uncompressed Block IO Status
			If CheckIO(StartPos - 1) Or CheckIO(StartPos - 1 - (LastByte - 1)) = 1 Then
				'If the block will be UIO than only (Lastbyte-1) bytes will fit,
				'So we only need to check that many bytes
				Buffer(AdLoPos - 1) = 0 'Set IO Flag
				AdHiPos = AdLoPos - 2   'Updae AdHiPos
				LastByte = AdHiPos - 1  'Update LastByte
			ElseIf CheckIO(StartPos - 1 - LastByte) = 1 Then
				'If only the last byte is UIO then this byte will be ignored
				'And one less bytes will be stored uncompressed
				AdHiPos = AdLoPos - 1   'IO flag is not set, update AdHiPos
				LastByte = AdHiPos - 2  'But LastByte is decreased by an additional value
				'As the very last byte would go UIO and would need an additional IO Flag byte
			Else
				'Block will not go UIO
				'IO Flag will not be set
				AdHiPos = AdLoPos - 1   'Update AdHiPos
				LastByte = AdHiPos - 1  'And LastByte
			End If

			SI = StartPos - LastByte                         'Update POffset

			Buffer(AdHiPos) = Int((PrgAdd + SI) / 256)  'SI is 1 based
			Buffer(AdLoPos) = (PrgAdd + SI) Mod 256     'SI is 1 based

			For I As Integer = 0 To LastByte - 1            '-1 because the first byte of the buffer is the bitstream
				Buffer(LastByte - I) = Prg(StartPos - I)
			Next

			Buffer(0) = &H80                                        'Set Copression Bit to 1 (=Uncompressed block)
			ByteCnt = 1

		End If

		BlockCnt += 1
		BufferCnt += 1
		UpdateByteStream()

		ResetBuffer()                       'Resets buffer variables

		NextFileInBuffer = False            'Reset Next File flag

		If SI < 0 Then Exit Sub             'We have reached the end of the file -> exit

		'If we have not reached the end of the file, then update buffer

		Buffer(ByteCnt) = (PrgAdd + SI) Mod 256
		AdLoPos = ByteCnt

		BlockUnderIO = CheckIO(SI)          'Check if last byte of prg could go under IO

		If BlockUnderIO = 1 Then
			ByteCnt -= 1
		End If

		Buffer(ByteCnt - 1) = Int((PrgAdd + SI) / 256) Mod 256
		AdHiPos = ByteCnt - 1
		ByteCnt -= 2
		LastByte = ByteCnt               'LastByte = the first byte of the ByteStream after and Address Bytes (253 or 252 with blockCnt)

		If (BlockCnt = 1) Or ((Seq(SI).Bit + 8 < LastByte * 8) And (LastFileOfPart = True)) Then
			'For the 2nd and last block only recalculate the first byte's sequence
			CalcBestSequence(IIf(SI > 1, SI, 1), IIf(SI > 1, SI, 1))
		Else
			'For all other blocks recalculate the first 256 bytes' sequence (max offset=256)
			CalcBestSequence(IIf(SI > 1, SI, 1), IIf(SI - MaxOffset > 1, SI - MaxOffset, 1))
		End If

		StartPos = SI

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

	Public Function ClosePart(Optional NextFileIO As Integer = 1, Optional LastPartOnDisk As Boolean = False) As Boolean
		On Error GoTo Err

		ClosePart = True

		If NewBlock = True Then GoTo NewB   'The part will start in a new block

		'ADDS NEW PART TAG (Long Match Tag + End Tag) TO THE END OF THE PART, AND RESERVES LAST BYTE IN BUFFER FOR BLOCK COUNT

		Dim Bytes As Integer = 6 'BYTES NEEDED: BlockCnt + Long Match Tag + End Tag + AdLo + AdHi + 1st Literal (ByC=7 if BlockUnderIO=true - checked at SequenceFits)
		'If LastPartOnDisk = True Then Bytes += 1
		Dim Bits As Integer = IIf((LitCnt = -1) Or (LitCnt Mod (MaxLitLen + 1) = MaxLitLen), 1, 0)   'Calculate whether Match Bit is needed for new part

		If SequenceFits(Bytes, Bits, NextFileIO) Then       'This will add the EndTag to the needed bytes

			'Buffer has enough space for New Part Tag and New Part Info and first Literal byte (and IO flag if needed)

			If Bits = 1 Then AddRBits(0, 1)

NextPart:   'Match Bit is not needed if this is the beginning of the next block
			FilesInBuffer += 1  'There is going to be more than 1 file in the buffer

			If (BufferCnt > 0) And (FilesInBuffer = 2) Then         'Reserve last byte in buffer for Block Count...
				For I = ByteCnt + 1 To 255                          '... only once, when the 2nd file is added to the same buffer
					Buffer(I - 1) = Buffer(I)
				Next
				ByteCnt -= 1
				Buffer(255) = 1                                     'Last byte reserved for BlockCnt
			End If

			'MsgBox(Hex(Buffer(BitCnt)))

			Buffer(ByteCnt) = LongMatchTag                          'Then add New File Match Tag
			Buffer(ByteCnt - 1) = EndTag
			ByteCnt -= 2

			If LastPartOnDisk = True Then       'This will finish the disk
				Buffer(ByteCnt) = ByteCnt - 2   'Finish disk with a dummy literal byte that overwrites itself to reset LastX for next disk side
				Buffer(ByteCnt - 1) = &H3       'New address is the next byte in buffer
				Buffer(ByteCnt - 2) = &H0       'Dummy $00 Literal that overwrites itself
				LitCnt = 0                      'One (dummy) literal
				'AddLitBits()                   'NOT NEEDED, WE ARE IN THE MIDDLE OF THE BUFFER, 1ST BIT NEEDS TO BE OMITTED
				AddRBits(0, 1)                  'ADD 2ND BIT SEPARATELY (0-BIT, TECHNCALLY, THIS IS NOT NEEDED SINCE THIS IS THE LAST BIT)
				'-------------------------------------------------------------------
				'Buffer(ByteCnt - 3) = &H0      'THIS IS THE END TAG, NOT NEEDED HERE, WILL BE ADDED WHEN BUFFER IS CLOSED
				'ByteCnt -= 4					'*BUGFIX, THANKS TO RAISTLIN FOR REPORTING
				'-------------------------------------------------------------------
				ByteCnt -= 3
			End If

			'DO NOT CLOSE LAST BUFFER HERE, WE ARE GOING TO ADD NEXT PART TO LAST BUFFER

			If ByteSt.Count > BlockPtr Then     'Only save block count if block is already added to ByteSt
				ByteSt(BlockPtr) = LastBlockCnt
				LoaderParts += 1
			End If

			LitCnt = -1                                                 'Reset LitCnt here
		Else
NewB:          'Next File Info does not fit, so close buffer
			CloseBuffer()               'Adds EndTag and starts new buffer
			'Then add 1 dummy literal byte to new block (blocks must start with 1 literal, next part tag is a match tag)
			Buffer(255) = &HFC          'Dummy Address ($03fc* - first literal's address in buffer... (*NextPart above, will reserve BlockCnt)
			Buffer(254) = &H3           '...we are overwriting it with the same value
			Buffer(253) = &H0           'Dummy value, will be overwritten with itself
			LitCnt = 0
			AddLitBits()                'WE NEED THIS HERE, AS THIS IS THE BEGINNING OF THE BUFFER, AND 1ST BIT WILL BE CHANGED TO COMPRESSION BIT
			ByteCnt = 252
			LastBlockCnt += 1

			If LastBlockCnt > 255 Then
				'Parts cannot be larger than 255 blocks compressed
				'There is some confusion here how PartCnt is used in the Editor and during Disk building...
				MsgBox("Part " + IIf(CompressPartFromEditor = True, PartCnt + 1, PartCnt).ToString + " would need " + LastBlockCnt.ToString + " blocks on the disk." + vbNewLine + vbNewLine + "Parts cannot be larger than 255 blocks!", vbOKOnly + vbCritical, "Part exceeds 255-block limit!")
				If CompressPartFromEditor = False Then GoTo NoGo
			End If

			BlockCnt -= 1
			'THEN GOTO NEXT PART SECTION
			GoTo NextPart
		End If

		NewBlock = SetNewBlock        'NewBlock is true at closing the previous part, so first it just sets NewBlock2
		SetNewBlock = False            'And NewBlock2 will fire at the desired part

		Exit Function
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

NoGo:
		ClosePart = False

	End Function

	Public Sub CloseFile()
		On Error GoTo Err

		'ADDS NEXT FILE TAG TO BUFFER

		'4 bytes and 0-1 bits needed for NextFileTag, Address Bytes and first Lit byte (+1 more if UIO)
		Dim Bytes As Integer = 4 'BYTES NEEDED: End Tag + AdLo + AdHi + 1st Literal (ByC=5 only if BlockUnderIO=true - checked at SequenceFits()
		Dim Bits As Integer = IIf((LitCnt = -1) Or (LitCnt Mod (MaxLitLen + 1) = MaxLitLen), 1, 0)   'Calculate whether Match Bit is needed for new file

		If SequenceFits(Bytes, Bits, CheckIO(PrgLen - 1)) Then

			'Buffer has enough space for New File Match Tag and New File Info and first Literal byte (and IO flag if needed)

			If Bits = 1 Then AddRBits(0, 1)

			Buffer(ByteCnt) = NextFileTag                           'Then add New File Match Tag
			ByteCnt -= 1
		Else
			'Next File Info does not fit, so close buffer
			CloseBuffer()
		End If

		Exit Sub
Err:
		MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

	End Sub

End Module
