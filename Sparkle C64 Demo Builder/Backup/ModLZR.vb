Friend Module ModLZR
	Dim CMStart, POffset, MLen As Integer
	Private ReadOnly MaxLongLen As Integer = 254
	Private ReadOnly MaxFarLen As Integer = 62
	Private ReadOnly MaxNearLen As Integer = 3
	Private ReadOnly DoDebug As Boolean = False
	Private LastByteCt As Integer = 255
	Private LastPOffset As Integer
	Private LastMOffset As Integer = 0
	Private LastMLen As Integer = 0

	Public Sub LZR()

		NM = 0 : FM = 0 : LM = 0

		MaxLit = 1 + 4 + 8 + 32 - 1             '=44

		ReDim LC(MaxLit)

		PrgAdd = Prg(1) * 256 + Prg(0)

		'Remove first 2 bytes from prg (AdLo and AdHi)
		For I As Integer = 0 To Prg.Length - 3
			Prg(I) = Prg(I + 2)
		Next

		ReDim Preserve Prg(Prg.Length - 3)

		PrgLen = Prg.Length

		OrigBC = Int(PrgLen / 256)
		If PrgLen Mod 256 <> 0 Then OrigBC += 1

		ReDim ByteSt(0)
		ReDim Buffer(255)

		'Initialize variables

		Buffer(0) = (PrgAdd + PrgLen - 1) Mod 256
		Buffer(1) = Int((PrgAdd + PrgLen - 1) / 256)
		BitCt = 2

		If BlockCnt = 0 Then
			ByteCt = 254         'First file in part, last byte of first block will be BlockCnt
		Else
			ByteCt = 255         'All other blocks
		End If

		LastByte = ByteCt
		BitPos = 15

		AddRBits(IOBit, 1)       'First bit in bitstream = IO status
		BufferCt = 0
		MatchStart = PrgLen - 1    '2
		LitCnt = 0
		Buffer(ByteCt) = Prg(PrgLen - 1)
		ByteCt -= 1

		For POffset = PrgLen - 2 To 0 Step -1       'Skip Last byte

Restart:
			If FindMatch() = True Then

				'MATCH
				'Longest Match found, now check if it is a Longmatch, a Farmatch or a Nearmatch
				If MaxLen > MaxFarLen Then
					'MaxLen=63-2557
					If LongMatch() = False Then GoTo Restart
				ElseIf MaxLen > MaxNearLen Then
					'MaxLen=4-62
					If FarMatch() = False Then GoTo Restart
				Else
					'MaxLen=1-3
					If MaxOffset > 64 Then
						'MaxOffset>64
						If FarMatch() = False Then GoTo Restart
					Else
						'MaxOffset=<64
						If NearMatch() = False Then GoTo Restart
					End If
				End If
			Else
				If Literal() = False Then GoTo Restart
			End If
		Next

		'Done with file, check literal counter 
		If LitCnt > -1 Then
			'And update bitstream if LitCnt > -1
			AddLitBits()
		End If

		'Check buffer
		'If (BitCt <> 2) Or If((BufferCt = 0) And (BlockCnt = 0), (ByteCt <> 254), (ByteCt <> 255)) Or (BitPos <> 15) Then
		'If (ByteCt <> 2) Or If(BufferCt = 0, (BitCt <> 254), (BitCt <> 255)) Or (BitPos <> 15) Then
		'If ByteCt <> LastByte Then
		'And save if not empty
		CloseBuff()
		'MsgBox(Hex(LastByte - ByteCt))
		'End If

		'ByteSt(255) = BufferCt 'this is done once, after the last file in part

		ReDim Prg(ByteSt.Count - 1)
		For I = 0 To ByteSt.Count - 1
			Prg(I) = ByteSt(I)
		Next

		BlockCnt += BufferCt

	End Sub

	Private Function FindMatch() As Boolean
		If DoDebug Then Debug.Print("FindMatch")
		'CMStart = If(POffset < MatchStart - 255, POffset + 255, MatchStart)     '255? vs 256?

		CMStart = POffset + 256

		If CMStart > MatchStart Then
			CMStart = MatchStart
		End If

		MatchCnt = 0
		ReDim MatchOffset(0), MatchLen(0), MatchSave(0)

		MLen = -1
		For CmPos = CMStart To POffset + 1 Step -1
NextByte:
			MLen += 1

			If POffset - MLen > -1 Then
				If Prg(POffset - MLen) = Prg(CmPos - MLen) And (MLen < MaxLongLen + 1) = True Then GoTo NextByte
			End If

				'Calculate length of matches
				If MLen > 1 Then
				MLen -= 1
				If (CmPos - POffset > 64) And (MLen < 2) Then
					'Farmatch, but too short (2 bytes only)
					'Ignore it
				Else
					'Save it to Match List
					MatchCnt += 1
					ReDim Preserve MatchOffset(MatchCnt)
					ReDim Preserve MatchLen(MatchCnt)
					ReDim Preserve MatchSave(MatchCnt)
					MatchOffset(MatchCnt) = CmPos - POffset 'Save offset
					MatchLen(MatchCnt) = MLen               'Save len

					If MLen > MaxFarLen Then                       'Calculate and save MatchSave
						'LongMatch (MLen=63-255)
						MatchSave(MatchCnt) = MLen - 2
					ElseIf MLen > MaxNearLen Then
						'FarMatch (MLen=04-62)
						MatchSave(MatchCnt) = MLen - 1
					Else    'MLen=00-03
						If MatchOffset(MatchCnt) > 64 Then
							'FarMatch
							MatchSave(MatchCnt) = MLen - 1
						Else
							'NearMatch
							MatchSave(MatchCnt) = MLen - 0
						End If
					End If
				End If
			End If
			MLen = -1   'Reset MLen
		Next

		If MatchCnt > 0 Then
			'MATCH

			'Debug.Print(MatchCnt.ToString)

			'Find longest match in Match List
			MaxSave = MatchSave(1)
			MaxLen = MatchLen(1)
			MaxOffset = MatchOffset(1)
			Match = 1
			For I = 1 To MatchCnt
				If MatchSave(I) = MaxSave Then

					If MatchLen(I) > MaxLen Then
						MaxSave = MatchSave(I)
						MaxLen = MatchLen(I)
						MaxOffset = MatchOffset(I)
						Match = I
					End If

					'THERE IS A BUG AND IT CAN BE FOUND USING THE CODE BELOW INSTEAD OF THE ONE ABOVE
					'BLOCK ADDRESS $6B84 (-$6929) IN SATURN - FIXED!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

					'If (MatchOffset(I) <= 64) And (MatchLen(I) <= MaxNearLen) Then         'Near Match, preferred
					'MaxSave = MatchSave(I)
					'MaxLen = MatchLen(I)
					'MaxOffset = MatchOffset(I)
					'Match = I
					'End If
				ElseIf MatchSave(I) > MaxSave Then  'Far Match, only update if saves more
					MaxSave = MatchSave(I)
					MaxLen = MatchLen(I)
					MaxOffset = MatchOffset(I)
					Match = I
				End If
			Next
			FindMatch = True
			If MaxSave = 0 Then
				FindMatch = False
			End If
		Else
			FindMatch = False
		End If

	End Function

	Private Function LongMatch() As Boolean
		If DoDebug Then Debug.Print("LongMatch")
		Dim Bits As Integer
		LongMatch = True
		'LONGMATCH
		'If MaxLen > MaxLongLen - 1 Then MaxLen = MaxLongLen - 1 'Cannot be longer than 254 bytes

		If LitCnt > -1 Then CheckShortMatch()

		If LitCnt > -1 Then 'there is a Literal Sequence to be finished first

			Select Case LitCnt
				Case 0
					Bits = 2 + 0
				Case 1 To 4
					Bits = 3 + 2 + 0
				Case 5 To 12
					Bits = 4 + 3 + 0
				Case 13 To MaxLit
					Bits = 4 + 5 + 0
				Case MaxLit + 1
					Bits = 4 + 5 + 1
			End Select
			'Check if Far Match fits
			If DataFits(3, Bits) = True Then
				'Yes, we have space
				If LitCnt > -1 Then AddLitBits()        'First close last literal sequence
				AddLongM()           'Then add Long Match
			ElseIf DataFits(2, Bits) = True Then    'Long match does not fit, check if a far match fits
				If LitCnt > -1 Then AddLitBits()
				If MaxLen > MaxFarLen Then MaxLen = MaxFarLen                  'Adjust Len to MaxFarLen
				AddFarM()
				GoTo BufferFull2
			ElseIf (MaxOffset <= 64) And DataFits(1, Bits) = True Then
				If LitCnt > -1 Then AddLitBits()
				If MaxLen > MaxNearLen Then MaxLen = MaxNearLen
				AddNearM()
				GoTo BufferFull2
			Else
				'Match does not fit, check if we can add byte as literal
				Select Case LitCnt + 1
					Case 0              'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
						Bits = 2
					Case 1 To 4
						Bits = 3 + 2
					Case 5 To 12
						Bits = 4 + 3
					Case 13 To MaxLit
						Bits = 4 + 5
					Case MaxLit + 1
						Bits = 4 + 5 + 2
				End Select
				If DataFits(1, Bits) Then
					AddLitByte()        'Add Byte As literal And update LitCnt
					GoTo BufferFull     'Then close buffer
				Else
					'Nothing fits, close buffer
BufferFull:
					AddLitBits()
BufferFull2:
					CloseBuff()
					LongMatch = False    'Goto Restart
				End If
			End If
		Else    'No literals, we need match tag (or literal tag)
			If DataFits(3, 1) = True Then
				'check if long match fits
				AddLongM()
			ElseIf DataFits(2, 1) = True Then
				'no, check if first 63 bytes fit a far match
				If MaxLen > MaxFarLen Then MaxLen = MaxFarLen
				AddFarM()       'Then add Far Bytes
				GoTo BufferFull2
			ElseIf (MaxOffset < 65) And DataFits(1, 1) = True Then
				'No, check if offset<64 and if first 3 bytes fit a near match
				If MaxLen > MaxNearLen Then MaxLen = MaxNearLen
				AddNearM()
				GoTo BufferFull2
			ElseIf DataFits(1, 2) Then    'This is the first literal
				'Match does not fit, check if we can add byte as literal
				AddLitByte()           'Add byte as literal and update LitCnt
				GoTo BufferFull     'Then close buffer
			Else
				'Nothing fits, close buffer (no literals)
				GoTo BufferFull2
			End If
		End If

		'			If DataFits(1, Bits) Then
		'			AddLitByte()        'Add Byte As literal And update LitCnt
		'			Else
		'			POffset += 1
		'		End If
		'			'Nothing fits, close buffer
		'			AddLitBits()
		'BufferFull2:
		'			CloseBuff()
		'			LongMatch = False    'Goto Restart
		'		End If
		'Else    'No literals, we need match tag (or literal tag)
		'If DataFits(3, 1) = True Then
		''check if long match fits
		'AddLongM()
		'ElseIf DataFits(2, 1) = True Then
		''no, check if first 63 bytes fit a far match
		'MaxLen = MaxFarLen
		'AddFarM()       'Then add Far Bytes
		'GoTo BufferFull2
		'ElseIf MaxOffset < 65 Then
		''No, check of offset<64
		'If DataFits(1, 1) = True Then
		''Check if first 3 bytes fit a near match
		'MaxLen = MaxNearLen
		'AddNearM()
		'GoTo BufferFull2
		'End If
		'ElseIf DataFits(1, 2) Then    'This is the first literal
		''Match does not fit, check if we can add byte as literal
		'AddLitByte()           'Add byte as literal and update LitCnt
		'GoTo BufferFull     'Then close buffer
		'Else
		''Nothing fits, close buffer (no literals)
		'GoTo BufferFull2
		'End If
		'End If

	End Function

	Private Function FarMatch() As Boolean
		If DoDebug Then Debug.Print("FarMatch")
		Dim Bits As Integer
		FarMatch = True
		'FARMATCH
		'If MaxLen > MaxFarLen - 1 Then MaxLen = MaxFarLen - 1 'Cannot be longer than 64 bytes

		'If POffset = 2346 Then MsgBox(Hex(MaxLen) + vbNewLine + Hex(MaxOffset))

		If LitCnt > -1 Then CheckShortMatch()

		If LitCnt > -1 Then 'there is a Literal Sequence to be finished first

			Select Case LitCnt
				Case 0
					Bits = 2 + 0
				Case 1 To 4
					Bits = 3 + 2 + 0
				Case 5 To 12
					Bits = 4 + 3 + 0
				Case 13 To MaxLit
					Bits = 4 + 5 + 0
				Case MaxLit + 1
					Bits = 4 + 5 + 1
			End Select
			'Check if Far Match fits
			If DataFits(2, Bits) = True Then
				'Yes, we have space
				AddLitBits()        'First close last literal sequence
				AddFarM()           'Then add Far Bytes
			ElseIf (MaxOffset <= 64) And DataFits(1, Bits) = True Then
				If LitCnt > -1 Then AddLitBits()
				If MaxLen > MaxNearLen Then MaxLen = MaxNearLen
				AddNearM()
				GoTo BufferFull2
			Else
				'Match does not fit, check if we can add byte as literal
				Select Case LitCnt + 1
					Case 0              'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
						Bits = 2
					Case 1 To 4
						Bits = 3 + 2
					Case 5 To 12
						Bits = 4 + 3
					Case 13 To MaxLit
						Bits = 4 + 5
					Case MaxLit + 1
						Bits = 4 + 5 + 2
				End Select
				If DataFits(1, Bits) Then
					AddLitByte()        'Add Byte As literal And update LitCnt
					GoTo BufferFull     'Then close buffer
				Else
					'Nothing fits, close buffer
BufferFull:
					AddLitBits()
BufferFull2:
					CloseBuff()
					FarMatch = False    'Goto Restart
				End If
			End If
		Else    'No literals, we need match tag (or literal tag)
			If DataFits(2, 1) = True Then
				'Yes, we have space
				AddFarM()       'Then add Far Bytes
			ElseIf (MaxOffset < 65) And DataFits(1, 1) = True Then
				'No, check if offset<64 and if first 3 bytes fit a near match
				If MaxLen > MaxNearLen Then MaxLen = MaxNearLen
				AddNearM()
				GoTo BufferFull2
			ElseIf DataFits(1, 2) Then    'This is the first literal
				'Match does not fit, check if we can add byte as literal
				AddLitByte()           'Add byte as literal and update LitCnt
				GoTo BufferFull     'Then close buffer
			Else
				'Nothing fits, close buffer (no literals)
				GoTo BufferFull2
			End If
		End If

	End Function

	Private Function NearMatch() As Boolean
		If DoDebug Then Debug.Print("NearMatch")
		NearMatch = True

		If LitCnt > -1 Then CheckShortMatch()

		If LitCnt > -1 Then

			Select Case LitCnt
				Case 0
					Bits = 2 + 0
				Case 1 To 4
					Bits = 3 + 2 + 0
				Case 5 To 12
					Bits = 4 + 3 + 0
				Case 13 To MaxLit
					Bits = 4 + 5 + 0
				Case MaxLit + 1
					Bits = 4 + 5 + 1
			End Select
			'we have already reserved space for literal bits, check is match fits
			If DataFits(1, Bits) = True Then
				'Yes, we have space
				AddLitBits()    'First close last literal sequence
				AddNearM()       'Then add Near Byte
			Else
				'Match does not fit, check if we can add byte as literal
				Select Case LitCnt + 1
					Case 0              'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
						Bits = 2
					Case 1 To 4
						Bits = 3 + 2
					Case 5 To 12
						Bits = 4 + 3
					Case 13 To MaxLit
						Bits = 4 + 5
					Case MaxLit + 1
						Bits = 4 + 5 + 2
				End Select
				If DataFits(1, Bits) Then
					'Match does not fit, check if we can add byte as literal
					AddLitByte()           'Add byte as literal and update LitCnt
					GoTo BufferFull     'Then close buffer
				Else
					'Nothing fits, close buffer
BufferFull:
					AddLitBits()
BufferFull2:
					CloseBuff()
					NearMatch = False    'Goto Restart
				End If
			End If
		Else
			If DataFits(1, 1) = True Then
				'Yes, we have space
				AddNearM()       'Then add Near Byte
			Else
				'If does not fit as Near Match, do not attempt to add as literal, as this would be the first literal, and would need more space
				'then as a match (1 byte + 1 bits as match) vs (1 byte + 2 bits as literal)
				'Nothing fits, close buffer (no literals)
				GoTo BufferFull2
			End If
		End If

	End Function

	Private Function Literal() As Boolean
		If DoDebug Then Debug.Print("Literal")
		Literal = True
		'LITERAL

		Select Case LitCnt + 1
			Case 0              'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
				Bits = 2
			Case 1 To 4
				Bits = 3 + 2
			Case 5 To 12
				Bits = 4 + 3
			Case 13 To MaxLit
				Bits = 4 + 5
			Case MaxLit + 1
				Bits = 4 + 5 + 2
		End Select

		If DataFits(1, Bits) = True Then 'We are adding 1 literal byte and 1+5 bits
			'We have space in the buffer for this byte
			AddLitByte()
		Else    'No space for this literal byte, close sequence
			If LitCnt > -1 Then
				AddLitBits()
BufferFull2:
				CloseBuff()
				Literal = False    'Goto Restart
			Else
				GoTo BufferFull2
			End If
		End If

	End Function

	Private Function DataFits(ByteLen As Integer, BitLen As Integer) As Boolean
		If DoDebug Then Debug.Print("DataFits")
		'Do not include closing bits in function call!!!

		Dim NeededBytes As Integer = 1  'Close Byte
		Dim CloseBit As Integer = 0     'CloseBit length - can overlap with Close Byte!!!

		If BitPos < 8 + BitLen + CloseBit Then '0 bits for Close Byte Match
			NeededBytes = 2     'Need an extra byte for bit sequence (closing match tag)
		End If

		If ByteCt >= BitCt + ByteLen + NeededBytes Then  '> because #$00 End Byte NEEDED
			DataFits = True
		Else
			DataFits = False
		End If

	End Function

	Private Sub AddLongM()
		If DoDebug Then Debug.Print("AddLongM")

		If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		SaveLastMatch()

		Buffer(ByteCt) = &HFC                   'Long Match Flag
		Buffer(ByteCt - 1) = MaxLen
		Buffer(ByteCt - 2) = MaxOffset - 1
		ByteCt -= 3

		POffset -= MaxLen
		LM += 1

		ResetLit()

	End Sub

	Private Sub AddFarM()
		If DoDebug Then Debug.Print("AddFarM")

		If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		SaveLastMatch()

		Buffer(ByteCt) = MaxLen * 4                         'Length of match (#$02-#$3f, cannot be #$00 (end byte), and #$01 - distant selector??)
		Buffer(ByteCt - 1) = MaxOffset - 1
		ByteCt -= 2

		POffset -= MaxLen
		FM += 1

		ResetLit()

	End Sub

	Private Sub AddNearM()
		If DoDebug Then Debug.Print("AddNearM")

		If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		SaveLastMatch()

		Buffer(ByteCt) = ((MaxOffset - 1) * 4) + MaxLen
		ByteCt -= 1

		POffset -= MaxLen
		NM += 1

		ResetLit()

	End Sub

	Private Sub AddLitByte()
		If DoDebug Then Debug.Print("AddLitByte")
		'Update bitstream

		If LitCnt = MaxLit Then 'If LitCnt=Max
			AddLitBits()        'Then update bitstream
			LitCnt = -1         'And start new literal sequence
			OverMaxLit = True
		End If
		LitCnt += 1             'Increase Literal Counter

		Buffer(ByteCt) = Prg(POffset)   'Add byte to Byte Stream
		ByteCt -= 1                     'Update Byte Position Counter

	End Sub

	Private Sub CheckOverlap()
		Dim PO As Integer = POffset + 1
		Dim MO, ML As Integer
		'CheckOverlap = True

		For I As Integer = PO + 1 To MatchStart - 1
			If I < PO + 256 Then
				'If current byte= byte(I) and previous byte (last byte of last match)=byte(I+1) then we have an overlap
				If (Prg(PO) = Prg(I)) And (Prg(PO + 1) = Prg(I + 1)) Then
					'Debug.Print("Overlap found:" + vbTab + "POffset: " + Hex(PO) + vbTab + "MOffset: " + Hex(I) + vbTab + "LastMLen: " + Hex(LastMLen) + vbTab +
					'			"LastMOffset:" + Hex(LastMOffset) + vbTab + "Prg(PO):" + Hex(Prg(PO)) + vbTab + "Prg(PO+1):" + Hex(Prg(PO + 1)))
					LastMLen -= 1   'Last Match length will be 1 less as 
					'Update Last Match
					If I < PO + 65 Then                        'Overlap is a shot match
						If DataFits(1, 1) Then  'Check if new 2-byte match will fit as a short match
							If (LastMOffset <= 64) And (LastMLen <= MaxNearLen) Then   'Update previous match as a near match
								Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
								ByteCt = LastByteCt - 1
							ElseIf LastMLen <= MaxFarLen Then                           'Update previous match as a far match
								Buffer(LastByteCt) = LastMLen * 4
								Buffer(LastByteCt - 1) = LastMOffset - 1
								ByteCt = LastByteCt - 2
							Else                                                        'Update previous match as a long match
								Buffer(LastByteCt) = &HFC
								Buffer(LastByteCt - 1) = LastMLen
								Buffer(LastByteCt - 2) = LastMOffset - 1
								ByteCt = LastByteCt - 3
							End If
							MO = MaxOffset      'Save next match's offset
							ML = MaxLen         'and length
							MaxOffset = I - PO
							MaxLen = 1
							LitCnt = -1
							AddNearM()
							MaxOffset = MO
							MaxLen = ML
							POffset = PO - 1
						End If
					Else                                            'Overlap is a long match - THIS WOULD ONLY SAVE 1 BIT
						If DataFits(2, 1) Then
							If (LastMOffset <= 64) And (LastMLen <= MaxNearLen) Then            'Make it a near match
								Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
								ByteCt = LastByteCt - 1
							ElseIf LastMLen <= MaxFarLen Then                                   'Make it a far match
								Buffer(LastByteCt) = LastMLen * 4
								Buffer(LastByteCt - 1) = LastMOffset - 1
								ByteCt = LastByteCt - 2
							Else                                                                'Make it a long match
								Buffer(LastByteCt) = &HFC
								Buffer(LastByteCt - 1) = LastMLen
								Buffer(LastByteCt - 2) = LastMOffset - 1
								ByteCt = LastByteCt - 3
							End If
							MO = MaxOffset
							ML = MaxLen
							MaxOffset = I - PO
							MaxLen = 1
							LitCnt = -1
							AddFarM()
							MaxOffset = MO
							MaxLen = ML
							POffset = PO - 1
						End If
					End If
					LastMLen = 0                                   'Reset last match
					Exit Sub
				End If
			End If
		Next
		LastMLen = 0                                   'Reset last match

	End Sub

	Private Sub CheckShortMatch()
		Dim MO, ML As Integer

		Select Case LitCnt
			Case 0
				If OverMaxLit = True Then       'LitCnt=45 actually
					For I As Integer = POffset + 2 To MatchStart - 1
						If (Prg(POffset + 1) = Prg(I)) And (Prg(POffset + 2) = Prg(I + 1)) And (I < POffset + 1 + 256) Then
							MO = MaxOffset
							ML = MaxLen
							MaxOffset = I - (POffset + 1)
							MaxLen = 1
							LitCnt = 43             '= 44 + 1 - 2
							ByteCt += 2             'Step back 2 bytes
							POffset += MaxLen + 1   'Repositon POffset


							Buffer(BitCt) = 0       'Delete unused bit stream data
							BitCt -= 1              'Step back 8+1 bits
							BitPos += 1             'we have already added 9 bits and 45 (0-44) Literal bytes to the bit and byte streams
							If BitPos > 15 Then     'So we are now removing 8+1 bits from the bit stream, and also update the BitPos pointer
								BitPos -= 8
								Buffer(BitCt) = 0   'Delete unused bit steam data
								BitCt -= 1
							End If
							Buffer(BitCt) = Buffer(BitCt) And (&H100 - (2 ^ (BitPos - 7)))  'Delete bit stream

							AddLitBits()            'We are re-adding bits for 43 Literal bytes
							AddFarM()               'This will also recet litCnt and OverMaxLit
							SM2 += 1
							MaxOffset = MO
							MaxLen = ML
							POffset -= 1
							Exit Sub
						End If
					Next
				Else
					'1-byte long Literal before next match, check for overlaps here
					If LastMLen > 2 Then CheckOverlap()
				End If
			Case 1, 2, 5, 6, 13, 14                         'rangemax+1,rangemax+2
				For I As Integer = POffset + 2 To MatchStart - 1
					If (Prg(POffset + 1) = Prg(I)) And (Prg(POffset + 2) = Prg(I + 1)) And (I < POffset + 1 + 256) Then
						MO = MaxOffset
						ML = MaxLen
						MaxOffset = I - (POffset + 1)
						MaxLen = 1
						LitCnt -= 2
						ByteCt += 2
						POffset += MaxLen + 1
						AddLitBits()
						AddFarM()
						SM2 += 1
						MaxOffset = MO
						MaxLen = ML
						POffset -= 1
						Exit Sub
					End If
				Next
		End Select

	End Sub

	Private Sub AddLitBits()
		If DoDebug Then Debug.Print("AddLitBits")

		If LitCnt = -1 Then Exit Sub

		'AddLitToArray()

		AddRBits(1, 1)               'Add Literal Selector

		Select Case LitCnt 'Mod 92
			Case 0
				AddRBits(0, 1)              'Add Literal Length Selector 0	- read no more bits
			Case 1 To 4
				AddRBits(2, 2)              'Add Literal Length Selector 10 - read 2 more bits
				AddRBits(LitCnt - 1, 2)     'Add Literal Length: 00-03, 2 bits	-> 1000 00xx when read
			Case 5 To 12
				AddRBits(6, 3)              'Add Literal Length Selector 110 - read 3 more bits
				AddRBits(LitCnt - 5, 3)     'Add Literal Length: 00-07, 3 bits	-> 1000 1xxx when read
			Case 13 To MaxLit
				AddRBits(7, 3)              'Add Literal Length Selector 111 - read 4 more bits
				AddRBits(LitCnt - 13, 5)    'Add Literal Length: 00-1f, 5 bits	-> 101x xxxx when read
			Case Else
				MsgBox("LitCnt should not be more than MaxLit!!!",, LitCnt.ToString)  'Doubt this can happen
		End Select

		'DO NOT RESET LitCnt HERE!!!

	End Sub

	Private Sub AddRBits(Bit As Integer, BCnt As Byte)
		If DoDebug Then Debug.Print("AddRBits:" + vbTab + Hex(Bit).ToString + vbTab + BCnt.ToString)

		Bits = Buffer(BitCt) * 256  'Load last bitstream byte to upper byte of Bits

		Bit = Bit * (2 ^ (8 - BCnt)) 'Shift bits to the leftmost position in byte

		'Add bits to bitstream
		For I = 1 To BCnt  '1-5 bits to be added
			Bit = Bit * 2  'xxxxxx00 -> x-xxxxx000
			Bits += Int(Bit / 256) * 2 ^ BitPos
			BitPos -= 1
			Bit = Bit Mod 256
		Next

		Buffer(BitCt) = Int(Bits / 256) 'Save upper byte to buffer

		If BitPos < 8 Then  'Lower byte is affected, save it to buffer, too
			BitCt += 1      'Next byte in BitStream
			Buffer(BitCt) = Bits Mod 256
			BitPos += 8     'Reset bitpos counter
		End If

	End Sub

	Private Sub CloseBuff()
		If DoDebug Then Debug.Print("CloseBuff")

		'Close Byte Match Tag = 0, so it is not needed as it is the default value in the next bit position!!!
		'#$00 EOF Byte not needed, it is the default value in buffer

		Buffer(2) = Buffer(2) And &HBF                    'Delete Compression Bit (Default (i.e. compressed) value is 0)

		'MsgBox(Hex(POffset) + vbNewLine + Hex(MatchStart))

		If MatchStart - POffset < LastByte - 2 Then                    'If bytes compressed < 252/253 then
			If MatchStart >= LastByte - 3 Then
				For I As Integer = 0 To LastByte - 3                   'If data cannot be compressed then simply copy 251/252 bytes
					Buffer(255 - I) = Prg(MatchStart - I)              'First 2 bytes are AdLo & AdHi, last byte may be block Cnt, next to last Bits
				Next
				Buffer(0) = (PrgAdd + MatchStart - (LastByte - 2)) Mod 256  'Update AdLo and AdHi with (Forward Address+1)
				Buffer(1) = Int((PrgAdd + MatchStart - (LastByte - 2)) / 256)
				Buffer(2) = (Buffer(2) Or &H40) And &HC0                    'And Set Copression Bit to 1 (=Uncompressed block)
				POffset = MatchStart - (LastByte - 2)                       'Update POffset
			ElseIf MatchStart < LastByte - 2 Then
				ReDim Buffer(255)
				For I As Integer = 0 To MatchStart
					Buffer(3 + I) = Prg(I)
				Next
				Buffer(0) = (PrgAdd - 1) Mod 256            'Update AdLo and AdHi with (Forward Address - 1)
				Buffer(1) = Int((PrgAdd - 1) / 256)
				Buffer(2) = (Buffer(2) Or &H40) And &HC0    'And set Copression Bit to 1 (=Uncompressed block)
				If MatchStart < LastByte - 3 Then
					Buffer(LastByte) = MatchStart + 1       'If block is not full then save number of bytes+1
					Buffer(2) = Buffer(2) Or &H20           'And set Length Flag to 1 (=Short Block)
				End If
			End If
		End If

		BufferCt += 1

		UpdateByteStream()

		'MsgBox(Hex(POffset))

		If POffset < 0 Then Exit Sub 'POffset = 0

		ReDim Buffer(255)

		Buffer(0) = (PrgAdd + POffset) Mod 256
		Buffer(1) = (Int((PrgAdd + POffset) / 256)) Mod 256

		BitCt = 2
		ByteCt = 255
		LastByte = ByteCt
		BitPos = 15
		AddRBits(IOBit, 1)              'IO Status Bit (1=On, 0=Off)

		Buffer(ByteCt) = Prg(POffset)   'Copy First Lit Byte to Buffer
		ByteCt -= 1                     'Update Byte Pos Counter
		MatchStart = POffset            'Update Match Start Flag
		POffset -= 1                    'Next Byte in Prg
		If POffset < 0 Then POffset = 0 'Unless it is <0		NEEDED DO NOT DELETE!!!
		LitCnt = 0                      'Literal counter has 1 value

	End Sub

	Private Sub AddLitToArray()

		LC(LitCnt) += 1

	End Sub

	Private Sub ResetLit()

		LitCnt = -1
		OverMaxLit = False

	End Sub

	Private Sub SaveLastMatch()

		LastByteCt = ByteCt
		LastPOffset = POffset
		LastMOffset = MaxOffset
		LastMLen = MaxLen

	End Sub

End Module
