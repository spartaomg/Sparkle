Friend Module ModLZV
	Dim CMStart, POffset, MLen As Integer
	Dim LitLen As Integer = 64

	Public Sub LZV()

		MaxLit = 2 + 4 + 8 + 16 - 1 '=29

		PrgLen = Prg.Length

		OrigBC = Int(PrgLen / 256)
		If PrgLen Mod 256 <> 0 Then OrigBC += 1

		ReDim ByteSt(0)
		ReDim Buffer(255)

		'Initialize variables
		BuffAdd = Prg(1) * 256 + Prg(0)
		Buffer(0) = Prg(0)
		Buffer(1) = Prg(1)
		ByteCt = 2
		If BlockCnt = 0 Then
			BitCt = 254         'First file in part, last byte of first block will be BlockCnt
			'Distant = 0         'Reset distant match pointer
			'DistBase = Prg(0) + (Prg(1) * 256)
		Else
			BitCt = 255         'All other blocks
		End If
		LastByte = BitCt
		BitPos = 15

		AddBits(IOBit, 1)       'First bit in bitstream = IO status
		'AddBits(0, 1)          'Compression Bit - NOT NEEDED HERE, First Literal Selector Bit will be overwritten
		BufferCt = 0
		MatchStart = 2
		LitCnt = -1

		For POffset = 2 To PrgLen - 1       'Check rest of file
Restart:
			If FindMatch() = True Then
				'MATCH
				'Longest Match found, now check if it is a Farmatch or Nearmatch
				If (MaxOffset > 63) Or (MaxLen > 4) Then
					If FarMatch() = False Then GoTo Restart
				Else
					If NearMatch() = False Then GoTo Restart
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
		If (ByteCt <> 2) Or If((BufferCt = 0) And (BlockCnt = 0), (BitCt <> 254), (BitCt <> 255)) Or (BitPos <> 15) Then
			'And save if not empty
			CloseBuff()
		End If

		'ByteSt(255) = BufferCt 'this is done once, after the last file in part

		ReDim Prg(ByteSt.Count - 1)
		For I = 0 To ByteSt.Count - 1
			Prg(I) = ByteSt(I)
		Next

		BlockCnt += BufferCt

	End Sub

	Private Function FindMatch() As Boolean

		CmStart = If(POffset > MatchStart + 255, POffset - 256, MatchStart)     '255? vs 256?

		MatchCnt = 0
		ReDim MatchOffset(0), MatchLen(0), MatchSave(0)

		For CmPos = CmStart To POffset - 2
			If Prg(POffset) = Prg(CmPos) Then
				'Find all matches
				MLen = 1
NextByte:
				If (POffset + MLen < PrgLen) And (CmPos + MLen < POffset) Then
					If Prg(POffset + MLen) = Prg(CmPos + MLen) Then
						MLen += 1
						GoTo NextByte
					Else
						'Calculate length of matches
CheckMLen:
						If MLen > 1 Then
							If (POffset - CmPos > 63) And (MLen < 3) Then
								'Farmatch, but too short (2 bytes only)
								'Ignore it
							Else
								'Save it to Match List
								MatchCnt += 1
								ReDim Preserve MatchOffset(MatchCnt)
								ReDim Preserve MatchLen(MatchCnt)
								ReDim Preserve MatchSave(MatchCnt)
								MatchOffset(MatchCnt) = POffset - CmPos 'Save offset
								If MLen > 4 Then   'Only far match can be longer then 4 bytes
									'Farmatch
									If MLen > LitLen Then MLen = LitLen   'Max length of far match=64
									MatchSave(MatchCnt) = MLen - 2
								Else                'Match=1-4 bytes can be either near or far
									If MatchOffset(MatchCnt) > 63 Then
										'Farmatch
										MatchSave(MatchCnt) = MLen - 2
									Else
										'Nearmatch
										MatchSave(MatchCnt) = MLen - 1
									End If
								End If
								MatchLen(MatchCnt) = MLen - 1  'Convert and save relative len
							End If
						End If
					End If
				Else
					GoTo CheckMLen
				End If
			End If
		Next

		If MatchCnt > 0 Then
			'MATCH

			'Find longest match in Match List
			MaxSave = MatchSave(1)
			MaxLen = MatchLen(1)
			MaxOffset = MatchOffset(1)
			Match = 1
			For I = 1 To MatchCnt
				If MatchSave(I) >= MaxSave Then
					If MatchOffset(I) < 64 Then         'Near Match, preferred
						MaxSave = MatchSave(I)
						MaxLen = MatchLen(I)
						MaxOffset = MatchOffset(I)
						Match = I
					ElseIf MatchSave(I) > MaxSave Then  'Far Match, only update if saves more
						MaxSave = MatchSave(I)
						MaxLen = MatchLen(I)
						MaxOffset = MatchOffset(I)
						Match = I
					End If
				End If
			Next
			FindMatch = True
		Else
			FindMatch = False
		End If

	End Function

	Private Function FarMatch() As Boolean
		Dim Bits As Integer
		FarMatch = True
		'FARMATCH
		If MaxLen > LitLen - 1 Then MaxLen = LitLen - 1 'Cannot be longer than 128 bytes

		If LitCnt > -1 Then 'there is a Literal Sequence to be finished first
			Select Case LitCnt
				Case 0 To 1
					Bits = 1 + 2 + 1 + 0
				Case 2 To 5
					Bits = 1 + 2 + 2 + 0
				Case 6 To 13
					Bits = 1 + 2 + 3 + 0
				Case 14 To MaxLit
					Bits = 1 + 2 + 4 + 0
				Case MaxLit + 1
					Bits = 1 + 2 + 4 + 1  'Also needs Match Tag
			End Select
			'Check if Far Match fits
			If DataFits(2, Bits) = True Then
				'Yes, we have space
				AddLitBits()        'First close last literal sequence
				AddFarM()           'Then add Far Bytes
			Else
				'Match does not fit, check if we can add byte as literal
				Select Case LitCnt + 1
					Case 0 To 1             'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
						Bits = 1 + 2 + 1
					Case 2 To 5
						Bits = 1 + 2 + 2
					Case 6 To 13
						Bits = 1 + 2 + 3
					Case 14 To MaxLit
						Bits = 1 + 2 + 4
					Case MaxLit + 1     'LitCnt+1>MaxLit, we need to start a new counter
						Bits = 1 + 2 + 4 + 1 + 2 + 1    '92 Literals + 1 more
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
			ElseIf DataFits(1, 1 + 2 + 1) Then    'This is the first literal
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
		NearMatch = True

		'Nearmatch
		If MaxLen > 3 Then MaxLen = 3   'Match cannot be longer than 4 bytes

		If LitCnt > -1 Then
			Select Case LitCnt
				Case 0 To 1
					Bits = 1 + 2 + 1 + 0
				Case 2 To 5
					Bits = 1 + 2 + 2 + 0
				Case 6 To 13
					Bits = 1 + 2 + 3 + 0
				Case 14 To MaxLit
					Bits = 1 + 2 + 4 + 0
				Case MaxLit + 1
					Bits = 1 + 2 + 4 + 1  'Also needs Match Tag
			End Select
			'we have already reserved space for literal bits, check is match fits
			If DataFits(1, Bits) = True Then
				'Yes, we have space
				AddLitBits()    'First close last literal sequence
				AddNearM()       'Then add Near Byte
			Else
				'Match does not fit, check if we can add byte as literal
				Select Case LitCnt + 1
					Case 0 To 1             'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
						Bits = 1 + 2 + 1
					Case 2 To 5
						Bits = 1 + 2 + 2
					Case 6 To 13
						Bits = 1 + 2 + 3
					Case 14 To MaxLit
						Bits = 1 + 2 + 4
					Case MaxLit + 1     'LitCnt+1>MaxLit, we need to start a new counter
						Bits = 1 + 2 + 5 + 1 + 2 + 1    '92 Literals + 1 more
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
				'then as a match (1 byte + 2 bits as match) vs (1 byte + 5 bits as literal)
				'Nothing fits, close buffer (no literals)
				GoTo BufferFull2
			End If
		End If

	End Function

	Private Function Literal() As Boolean
		Literal = True
		'LITERAL

		Select Case LitCnt + 1
			Case 0 To 1             'Cannot be 0 (we already have literals and now we add 1 more, minimum is 0+1=1)
				Bits = 1 + 2 + 1
			Case 2 To 5
				Bits = 1 + 2 + 2
			Case 6 To 13
				Bits = 1 + 2 + 3
			Case 14 To MaxLit
				Bits = 1 + 2 + 4
			Case MaxLit + 1     'LitCnt+1>MaxLit, we need to start a new counter
				Bits = 1 + 2 + 5 + 1 + 2 + 1    '92 Literals + 1 more
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
		'Do not include closing bits in function call!!!

		Dim NeededBytes As Integer = 1

		If BitPos < 8 + BitLen + 2 Then '2 bits for Match=00
			NeededBytes = 2     'Need an extra byte for bit sequence (closing match tag)
		End If

		If BitCt >= ByteCt + ByteLen + NeededBytes Then  '> because #$00 End Byte NEEDED
			DataFits = True
		Else
			DataFits = False
		End If

	End Function

	Private Sub AddFarM()

		If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

		Buffer(ByteCt) = MaxLen * 4                         'Length of match (0-127)
		Buffer(ByteCt + 1) = (MaxOffset - 1) Xor &HFF   'To skip eor #$ff in code!
		ByteCt += 2
		'MsgBox(Hex(POffset + 8190).ToString + vbNewLine + Hex(MaxLen).ToString,, "AddFarM")
		POffset += MaxLen

		LitCnt = -1

	End Sub

	Private Sub AddNearM()

		If (LitCnt = -1) Or (LitCnt = MaxLit) Then
			AddBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag
		End If

		Buffer(ByteCt) = ((MaxOffset - 1) * 4) + MaxLen
		ByteCt += 1
		'MsgBox(Hex(POffset + 8190).ToString + vbNewLine + Hex(MaxLen).ToString,, "AddNearM")
		POffset += MaxLen

		LitCnt = -1

	End Sub

	Private Sub AddLitByte()
		'Update bitstream

		LitCnt += 1             'Increase Literal Counter
		If LitCnt > MaxLit Then 'If LitCnt > Max
			LitCnt = MaxLit
			AddLitBits()        'Then update bitstream, and start new literal byte stream
			LitCnt = 0          'This is the first byte of the new literal stream
		End If

		Buffer(ByteCt) = Prg(POffset)
		ByteCt += 1

	End Sub

	Private Sub AddLitBits()

		If LitCnt = -1 Then Exit Sub

		AddBits(1, 1)               'Add Literal Selector

		Select Case LitCnt 'Mod 92
			Case 0 To 1
				AddBits(0, 2)       'Add Literal Length Selector 00 - 2 more bits
				AddBits(LitCnt, 1)      'Add Literal Length: 00-01, 1 bit	-> 1000 001x when read
			Case 2 To 5
				AddBits(1, 2)       'Add Literal Length Selector 01 - 3 more bits
				AddBits(LitCnt - 2, 2)  'Add Literal Length: 00-03, 2 bits	-> 1000 01xx when read
			Case 6 To 13
				AddBits(2, 2)       'Add Literal Length Selector 10 - 4 more bits
				AddBits(LitCnt - 6, 3)  'Add Literal Length: 00-07, 3 bits	-> 1000 1xxx when read
			Case 14 To MaxLit
				AddBits(3, 2)       'Add Literal Length Selector 11 - 6 more bits
				AddBits(LitCnt - 14, 4) 'Add Literal Length: 00-0f, 4 bits	-> 1001 xxxx when read
			Case Else
				MsgBox("LitCnt should not be more than MaxLit!!!")  'Doubt this can happen
		End Select

		'DO NOT RESET LitCnt HERE!!!

	End Sub

	Private Sub CloseBuff()

		'MsgBox(Hex(POffset + 8190).ToString,, "CloseBuff")
		If (LitCnt = MaxLit) Or (LitCnt = -1) Then AddBits(0, 1)   '0 - 1 bit Match Tag only if needed
		'#$00 EOF Byte not needed, it is the default value in buffer

		Buffer(LastByte) = Buffer(LastByte) And &HBF                    'Delete Compression Bit (Default(compressed)=0)

		If POffset - MatchStart <= LastByte - 3 Then                    'If bytes compressed <= 251/252 then
			If MatchStart + LastByte - 3 < Prg.Length Then
				For I As Integer = 0 To LastByte - 3                    'If data cannot be compressed then simply copy 251/252 bytes
					Buffer(2 + I) = Prg(MatchStart + I)                 'First 2 bytes are AdLo & AdHi, last byte may be block Cnt, next to last Bits
				Next
				Buffer(LastByte) = (Buffer(LastByte) Or &H40) And &HC0  'and Set Copression Bit to 1 (=Uncompressed block)
				POffset = MatchStart + LastByte - 2
				BitCt = LastByte - 2
			End If
		End If

		'If (BlockCnt = 0) And (BufferCt = 0) Then       'If this is the first block in a file
		'Distant = POffset - 1                         'Save Data Pointer for distant match check
		'End If

		BufferCt += 1

		UpdateByteStream()

		ReDim Buffer(255)

		Buffer(0) = (BuffAdd + POffset - 2) Mod 256
		Buffer(1) = (Int((BuffAdd + POffset - 2) / 256)) Mod 256

		ByteCt = 2
		BitCt = 255
		LastByte = BitCt
		BitPos = 15
		AddBits(IOBit, 1)   'IO Status Bit (1=On, 0=Off)
		'AddBits(0, 1)       'Compression Bit - NOT NEEDED HERE, First Literal Selector Bit will be overwritten
		LitCnt = -1
		MatchStart = POffset

	End Sub

End Module
