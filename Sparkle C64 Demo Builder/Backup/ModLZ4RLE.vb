Friend Module ModLZ4RLE

	Public Sub LZ4RLE()

		MaxBit = 4
		MaxLit = (2 ^ MaxBit) - 1

		PrgLen = Prg.Length '.Count

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
		Else
			BitCt = 255         'All other blocks
		End If
		BitPos = 15
		AddBits(IOBit, 1)       'First bit in bitstream = IO status
		BufferCt = 0
		MatchStart = 2
		LitCnt = -1


		For DtPos = 2 To PrgLen - 1       'Check rest of file
Restart:

			RLECnt = 0
			'Look for RLE sequence
			For I = DtPos + 1 To PrgLen - 1
				If Prg(I) = Prg(DtPos) Then
					RLECnt += 1
				Else
					Exit For
				End If
			Next
			'Do we have at least 2 bytes in a row with the same value?
			If RLECnt > 1 Then    'RLE
				'MsgBox("RLE: " + RLECnt.ToString)
				'MsgBox(Hex(DtPos).ToString + vbNewLine + Hex(RLECnt).ToString)
				If LitCnt > -1 Then

					'We have already reserved space for literal bits, check is RLE fits
					If DataFits(2, 5 + 2) = True Then
						'Yes, we have space
						AddLitBits()    'First close last literal sequence
						AddRLE()       'Then add Near Byte
					ElseIf DataFits(1, 5) Then
						'RLE does not fit, check if we can add byte as literal
						AddLitByte()        'Add byte as literal and update LitCnt
						GoTo BufferFull     'Then close buffer
					Else
						'Nothing fits, close buffer
						GoTo BufferFull
					End If
				Else
					If DataFits(2, 2) = True Then
						AddRLE()

					ElseIf DataFits(1, 5) = True Then 'RLE does not fit, check if byte fits as a literal byte
						'We are adding 1 literal byte and 1+5 bits
						AddLitByte()
						'AddLitBits()
						DtPos += 1
						'CloseBuff()
						GoTo Restart
					Else    'No space for this literal byte, close sequence
						If LitCnt > -1 Then
							GoTo BufferFull
						Else
							GoTo BufferFull2
						End If
					End If
				End If
				LitCnt = -1
			Else    'No RLE, check for Match
				If DtPos > MatchStart + 255 Then
					CmLast = DtPos - 255
				Else
					CmLast = MatchStart
				End If

				MatchCnt = 0
				ReDim MatchOffset(0), MatchLen(0)

				For CmPos = CmLast To DtPos - 1
					If Prg(DtPos) = Prg(CmPos) Then
						'Find all matches
						DtLen = 1
NextByte:
						If (DtPos + DtLen < PrgLen) And (CmPos + DtLen < DtPos) Then
							If Prg(DtPos + DtLen) = Prg(CmPos + DtLen) Then
								DtLen += 1
								GoTo NextByte
							Else
								'Calculate length of matches
CheckDtLen:
								If DtLen > 1 Then
									If (DtPos - CmPos > 63) And (DtLen < 3) Then
										'Farmatch, but too short (2 bytes only)
										'Ignore it
									Else
										'Save it to Match List
										MatchCnt += 1
										ReDim Preserve MatchOffset(MatchCnt)
										ReDim Preserve MatchLen(MatchCnt)
										ReDim Preserve MatchSave(MatchCnt)
										MatchOffset(MatchCnt) = DtPos - CmPos 'Save offset
										If DtLen > 4 Then   'Only far match can be longer then 4 bytes
											'Farmatch
											If DtLen > 64 Then DtLen = 64   'Max length of far match=32
											MatchSave(MatchCnt) = DtLen - 2
										Else                'Match=1-4 bytes can be either near or far
											If MatchOffset(MatchCnt) > 63 Then
												'Farmatch
												MatchSave(MatchCnt) = DtLen - 2
											Else
												'Nearmatch
												MatchSave(MatchCnt) = DtLen - 1
											End If
										End If
										'If MatchOffset(MatchCnt) > 63 Then
										'Farmatch, len=0-63
										'If DtLen > 64 Then DtLen = 64
										'Saved bytes
										'MatchSave(MatchCnt) = DtLen - 2
										'Else
										'Nearmatch, len=0-3
										'If DtLen > 4 Then DtLen = 4
										'Saved bytes
										'MatchSave(MatchCnt) = DtLen - 1
										'End If
										MatchLen(MatchCnt) = DtLen - 1  'Convert and save len
									End If
								End If
							End If
						Else
							GoTo CheckDtLen
						End If
					End If
				Next

				'-------------------------------------------------------------------------------------------

				If MatchCnt > 0 Then
					'MATCH

					'Find longest match in Match List
					MaxSave = 0
					For I = 1 To MatchCnt
						If MatchSave(I) > MaxSave Then
							MaxSave = MatchSave(I)
							MaxLen = MatchLen(I)
							MaxOffset = MatchOffset(I)
							Match = I
						End If
					Next

					'Debug.Print(MaxSave.ToString)
					'MsgBox("Match: " + MaxLen.ToString)

					'Longest Match found, now check if it is a Farmatch or Nearmatch
					If MaxOffset > 63 Then
						'Farmatch, len>=3 by default
						If MaxLen > 63 Then MaxLen = 63 'Cannot be longer than 64 bytes

						If LitCnt > -1 Then
							'we have already reserved space for literal bits, check is match fits
							If DataFits(2, 5 + 2) = True Then
								'Yes, we have space
								AddLitBits()        'First close last literal sequence
								AddFarM()       'Then add Far Bytes
							ElseIf DataFits(1, 5) Then
								'Match does not fit, check if we can add byte as literal
								AddLitByte()           'Add byte as literal and update LitCnt
								GoTo BufferFull     'Then close buffer
							Else
								'Nothing fits, close buffer
BufferFull:
								AddLitBits()
BufferFull2:
								CloseBuff()
								GoTo Restart

							End If
						Else    'No literals, we need match tag (or literal tag)
							If DataFits(2, 2) = True Then
								'Yes, we have space
								AddFarM()       'Then add Far Bytes
							ElseIf DataFits(1, 5) Then    'This is the first literal
								'Match does not fit, check if we can add byte as literal
								AddLitByte()           'Add byte as literal and update LitCnt
								GoTo BufferFull     'Then close buffer
							Else
								'Nothing fits, close buffer (no literals)
								GoTo BufferFull2
							End If
						End If
					Else
						'Nearmatch
						If MaxLen > 3 Then MaxLen = 3   'Match cannot be longer than 4 bytes

						If LitCnt > -1 Then
							'we have already reserved space for literal bits, check is match fits
							If DataFits(1, 5 + 2) = True Then
								'Yes, we have space
								AddLitBits()    'First close last literal sequence
								AddNearM()       'Then add Near Byte
							ElseIf DataFits(1, 5) Then
								'Match does not fit, check if we can add byte as literal
								AddLitByte()           'Add byte as literal and update LitCnt
								GoTo BufferFull     'Then close buffer
							Else
								'Nothing fits, close buffer
								GoTo BufferFull
							End If
						Else
							If DataFits(1, 2) = True Then
								'Yes, we have space
								AddNearM()       'Then add Near Byte
							Else
								GoTo BufferFull2
							End If
						End If
					End If
					'Reset LitCnt
					LitCnt = -1
				Else
					'MsgBox("Literal: " + LitCnt.ToString)
					'NO MATCH
					'This is a literal Byte, add it to bytestream
					If DataFits(1, 5) = True Then 'We are adding 1 literal byte and 1+5 bits
						'We have space in the buffer for this byte
						AddLitByte()
					Else    'No space for this literal byte, close sequence
						If LitCnt > -1 Then
							GoTo BufferFull
						Else
							GoTo BufferFull2
						End If
					End If
				End If
			End If
		Next

		'Done with file, check literal counter 
		If LitCnt > -1 Then
			'And update bitstream if LitCnt > -1
			AddLitBits()
		End If

		'Check buffer
		If (ByteCt <> 2) Or IIf((BufferCt = 0) And (BlockCnt = 0), (BitCt <> 254), (BitCt <> 255)) Or (BitPos <> 15) Then
			'And save if not empty
			CloseBuff()
		End If

		ReDim Prg(ByteSt.Count - 1)
		For I = 0 To ByteSt.Count - 1
			Prg(I) = ByteSt(I)
		Next

		BlockCnt += BufferCt

	End Sub

	Private Sub AddRLE()

		AddBits(1, 2)   '01     - 2 bits RLE

		If RLECnt < 256 Then
			Buffer(ByteCt) = RLECnt
			Buffer(ByteCt + 1) = Prg(DtPos)
			ByteCt += 2
			DtPos += RLECnt
		Else
			Buffer(ByteCt) = 255
			Buffer(ByteCt + 1) = Prg(DtPos)
			ByteCt += 2
			DtPos += 255
		End If

	End Sub

	Private Function DataFits(ByteLen As Byte, BitLen As Byte) As Boolean
		'Do not include closing bits in function call!!!

		Dim NeededBytes As Byte

		NeededBytes = 1

		If BitPos < 8 + BitLen + 2 Then
			NeededBytes = 2     'Need an extra byte for bit sequence (closing match tag)
		End If

		If BitCt > ByteCt + ByteLen + NeededBytes Then  '> because #$00 End Byte NEEDED!!!!
			DataFits = True
		Else
			DataFits = False
		End If

	End Function

	Private Sub AddLitBits()

		If LitCnt = -1 Then Exit Sub
		AddBits(1, 1)           '1 - 1 bit
		AddBits(LitCnt, 4)      '4 more bits

	End Sub

	Private Sub AddLitByte()
		'Update bitstream
		If LitCnt < MaxLit Then 'If LitCnt < Max
			LitCnt += 1         'Then increase it
		Else                    'If LitCnt = Max
			AddLitBits()        'Then update bitstream, and start new literal byte stream
			LitCnt = 0          'This is the first byte of the new literal stream
		End If

		Buffer(ByteCt) = Prg(DtPos)
		ByteCt += 1

	End Sub

	Private Sub AddFarM()

		AddBits(0, 2)   '00     - 2 bits Match

		Buffer(ByteCt) = MaxLen * 4             'Length of match*4
		Buffer(ByteCt + 1) = (MaxOffset - 1) Xor &HFF   'To skip eor #$ff in code!
		ByteCt += 2
		DtPos += MaxLen

	End Sub

	Private Sub AddNearM()

		AddBits(0, 2)   '00     - 2 bits Match

		Buffer(ByteCt) = ((MaxOffset - 1) * 4) + MaxLen
		ByteCt += 1
		DtPos += MaxLen

	End Sub

	Private Sub CloseBuff()

		AddBits(0, 2)   '00 - 2 bits, Match tag
		'#$00 End Byte not needed, it is the default value in buffer

		BufferCt += 1

		UpdateByteStream()

		ReDim Buffer(255)   'Clear Buffer

		Buffer(0) = (BuffAdd + DtPos - 2) Mod 256
		Buffer(1) = (Int((BuffAdd + DtPos - 2) / 256)) Mod 256

		ByteCt = 2
		BitCt = 255

		BitPos = 15
		AddBits(IOBit, 1)

		LitCnt = -1
		MatchStart = DtPos

	End Sub

End Module
