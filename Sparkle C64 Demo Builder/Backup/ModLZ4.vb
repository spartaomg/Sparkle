Friend Module ModLZ4

	Public Sub LZ4()

		MaxBit = 4
		MaxLit = (2 ^ MaxBit) - 1

		PrgLen = Prg.Count

		OrigBC = Int(PrgLen / 256)
		If PrgLen Mod 256 <> 0 Then OrigBC += 1

		ReDim ByteSt(0)
		ReDim Buffer(255)

		'Initialize variables
		BuffAdd = Prg(1) * 256 + Prg(0)
		Buffer(0) = Prg(0)
		Buffer(1) = Prg(1)
		'MsgBox(Hex(BuffAdd))
		ByteCt = 2
		If BlockCnt = 0 Then
			BitCt = 254         'First file in part, last byte of first block will be BlockCnt
			Distant = 0         'Reset distant match pointer
			DistBase = Prg(0) + (Prg(1) * 256)
		Else
			BitCt = 255         'All other blocks
		End If
		LastByte = BitCt
		BitPos = 15

		AddBits(IOBit, 1)       'First bit in bitstream = IO status
		AddBits(0, 1)           'Compression bit
		BufferCt = 0
		MatchStart = 2
		LitCnt = -1

		For DtPos = 2 To PrgLen - 1       'Check rest of file
Restart:
			If DtPos > MatchStart + 255 Then
				CmLast = DtPos - 255
			Else
				CmLast = MatchStart
			End If

			MatchCnt = 0
			ReDim MatchOffset(0), MatchLen(0), MatchSave(0)

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
							'Calculate length of matches (DtLen = actual number of bytes)
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
									'Farmatch
									'If DtLen > 32 Then DtLen = 32
									'Saved bytes
									'MatchSave(MatchCnt) = DtLen - 2
									'Else
									'Nearmatch
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

			DistCnt = 0
			ReDim DistAd(0), DistLen(0), DistSave(0)

			If Distant <> 0 Then
				For I As Integer = 2 To Distant - 1
					If Prg(DtPos) = Prg(I) Then

						'Find all matches
						DtLen = 1
NextDist:
						If (DtPos + DtLen < PrgLen) And (I + DtLen < Distant) Then
							If Prg(DtPos + DtLen) = Prg(I + DtLen) Then
								DtLen += 1
								GoTo NextDist
							Else
								'Calculate length of matches (DtLen = actual number of bytes)
CheckDistLen:
								If DtLen > 1 Then
									If DtLen < 4 Then
										'Distantmatch, but too short (3 bytes only)
										'Ignore it
									Else
										'Save it to Match List
										DistCnt += 1
										ReDim Preserve DistAd(DistCnt)
										ReDim Preserve DistLen(DistCnt)
										ReDim Preserve DistSave(DistCnt)
										If DtLen > 32 Then DtLen = 32
										DistAd(DistCnt) = DistBase + I - 2 'Offset
										DistLen(DistCnt) = DtLen - 1  'Convert and save len
										DistSave(DistCnt) = DtLen - 3 'Saved bytes
									End If
								End If
							End If
						Else
							GoTo CheckDistLen
						End If

					End If
				Next
			End If
			DistCnt = 0
			If (MatchCnt > 0) Or (DistCnt > 0) Then
				'MATCH OR DISTANT MATCH

				MaxSave = 0
				'Find longest match in Match List

				If MatchCnt > 0 Then
					MaxSave = MatchSave(1)
					MaxLen = MatchLen(1)
					MaxOffset = MatchOffset(1)
					Match = 1
					MaxType = 1
					For I As Integer = 1 To MatchCnt
						If MatchOffset(I) < 64 Then       'Prefer nearmatch over farmatch
							If MatchSave(I) >= MaxSave Then
								MaxSave = MatchSave(I)
								MaxLen = MatchLen(I)
								MaxOffset = MatchOffset(I)
								Match = I
							End If
						ElseIf MatchSave(I) > MaxSave Then
							MaxSave = MatchSave(I)
							MaxLen = MatchLen(I)
							MaxOffset = MatchOffset(I)
							Match = I
							MaxType = 1
						End If
					Next
				End If

				'If DistCnt > 0 Then
				'If MaxSave = 0 Then
				'MaxSave = DistSave(1)
				'MaxLen = DistLen(1)
				'MaxOffset = DistAd(1)
				'Match = 1
				'MaxType = 2
				'End If
				'For I As Integer = 1 To DistCnt
				'If DistSave(I) > MaxSave Then
				'MaxSave = DistSave(I)
				'MaxLen = DistLen(I)
				'MaxOffset = DistAd(I)
				'Match = I
				'MaxType = 2
				'End If
				'Next
				'End If

				'Debug.Print(MaxSave.ToString)

				'MsgBox(MaxOffset.ToString)

				'Longest Match found, now check if it is a Farmatch or Nearmatch
				If MaxType = 1 Then     'FARMATCH OR NEARMATCH
					If (MaxOffset > 63) Or (MaxLen > 3) Then
						'Farmatch, len>=3 by default
						If MaxLen > 63 Then MaxLen = 63 'Cannot be longer than 64 bytes

						If LitCnt > -1 Then
							'we have already reserved space for 6 bits, check is match fits
							If CheckSpace(2, IIf(LitCnt < MaxLit, MaxBit + 1, MaxBit + 2)) = True Then
								'MsgBox("FarMatch")
								'Yes, we have space
								AddLiteralBits()    'First close last literal sequence
								AddFarMatch()       'Then add Far Bytes
							ElseIf CheckSpace(1, MaxBit + 1) Then
								'Match does not fit, check if we can add byte as literal
								AddByte()           'Add byte as literal and update LitCnt
								GoTo BufferFull     'Then close buffer
							Else
								'Nothing fits, close buffer
BufferFull:
								AddLiteralBits()
BufferFull2:
								CloseBuffer()
								GoTo Restart

							End If
						Else    'No literals, we need match tag (or literal tag)
							If CheckSpace(2, 1) = True Then
								'Yes, we have space
								'MsgBox("FarMatch, no Lit")
								AddFarMatch()       'Then add Far Bytes
							ElseIf CheckSpace(1, MaxBit + 1) Then
								'Match does not fit, check if we can add byte as literal
								AddByte()           'Add byte as literal and update LitCnt
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
							'we have already reserved space for 6 bits, check is match fits
							If CheckSpace(1, IIf(LitCnt < MaxLit, MaxBit + 1, MaxBit + 2)) = True Then
								'Yes, we have space
								AddLiteralBits()    'First close last literal sequence
								AddNearMatch()       'Then add Near Byte
							ElseIf CheckSpace(1, MaxBit + 1) Then
								'Match does not fit, check if we can add byte as literal
								AddByte()           'Add byte as literal and update LitCnt
								GoTo BufferFull     'Then close buffer
							Else
								'Nothing fits, close buffer
								GoTo BufferFull
							End If
						Else
							If CheckSpace(1, 1) = True Then
								'Yes, we have space
								AddNearMatch()       'Then add Near Byte
							Else
								GoTo BufferFull2
							End If
						End If
					End If
					'Reset LitCnt
					'LitCnt = -1
				Else    'DISTANT MATCH
					If MaxLen > 31 Then MaxLen = 31   'Match cannot be longer than 32 bytes

					If LitCnt > -1 Then
						'we have already reserved space for 6 bits, check is match fits
						If CheckSpace(1, IIf(LitCnt < MaxLit, MaxBit + 1, MaxBit + 2)) = True Then
							'Yes, we have space
							'MsgBox("DistantMatch, Lit")
							AddLiteralBits()    'First close last literal sequence
							AddDistantMatch()   'Then add Distant Bytes
						ElseIf CheckSpace(1, MaxBit + 1) Then
							'Match does not fit, check if we can add byte as literal
							AddByte()           'Add byte as literal and update LitCnt
							GoTo BufferFull     'Then close buffer
						Else
							'Nothing fits, close buffer
							GoTo BufferFull
						End If
					Else
						If CheckSpace(3, 1) = True Then
							'Yes, we have space
							'MsgBox("DistantMatch, no Lit")
							AddDistantMatch()       'Then add Far Bytes
						ElseIf CheckSpace(1, MaxBit + 1) Then
							'Match does not fit, check if we can add byte as literal
							AddByte()           'Add byte as literal and update LitCnt
							GoTo BufferFull     'Then close buffer
						Else
							'Nothing fits, close buffer (no literals)
							GoTo BufferFull2
						End If
					End If
				End If
				'Reset LitCnt
				LitCnt = -1
			Else
				'NO MATCH
				'This Is a literal Byte, add it to bytestream
				If CheckSpace(1, MaxBit + 1) = True Then 'We are adding 1 literal byte and 1+5 bits
					'We have space in the buffer for this byte
					AddByte()
				Else    'No space for this literal byte, close sequence
					If LitCnt > -1 Then
						GoTo BufferFull
					Else
						GoTo BufferFull2
					End If
				End If
			End If
		Next

		'Done with file, check literal counter 
		If LitCnt > -1 Then
			'And update bitstream if LitCnt > -1
			AddLiteralBits()
		End If

		'Check buffer
		If (ByteCt <> 2) Or IIf((BufferCt = 0) And (BlockCnt = 0), (BitCt <> 254), (BitCt <> 255)) Or (BitPos <> 15) Then
			'And save if not empty
			CloseBuffer()
		End If

		'ByteSt(255) = BufferCt 'this is done once, after the last file in part

		ReDim Prg(ByteSt.Count - 1)
		For I = 0 To ByteSt.Count - 1
			Prg(I) = ByteSt(I)
		Next

		BlockCnt += BufferCt

	End Sub

	Public Sub AddBits(Bit As Integer, BCnt As Byte)    'THIS IS ALSO USED BY LZ4+RLE!!!

		Bits = Buffer(BitCt) * 256  'Load last bitstream byte to upper byte of Bits

		Bit = Bit * (2 ^ (8 - BCnt)) 'Shift bits to the leftmost position in byte

		For I = 1 To BCnt  '2-6 bits to be added
			Bit = Bit * 2  'xxxxxx00 -> x-xxxxx000
			Bits += Int(Bit / 256) * 2 ^ BitPos
			BitPos -= 1
			Bit = Bit Mod 256
		Next

		Buffer(BitCt) = Int(Bits / 256) 'Save upper byte to buffer

		If BitPos < 8 Then  'Lower byte is affected, save it to buffer, too
			BitCt -= 1
			Buffer(BitCt) = Bits Mod 256
			BitPos += 8     'Reset bitpos counter
		End If

	End Sub

	Private Function CheckSpace(ByteLen As Byte, BitLen As Byte) As Boolean
		'Do not include closing byte and bit in function call!!!

		Dim NeededBytes As Byte
		Dim NeededBits As Byte

		NeededBytes = 1  'Need 1 extra byte for close byte

		'Check if we have pending literals
		If (LitCnt = -1) Or (LitCnt = MaxLit) Then
			'No pending literals, or LitCnt=31, need to save 1 bit for match tag
			NeededBits = 1
		Else
			'If LitCnt=0-30 then, next item must be a match, we do not need a match tag
			NeededBits = 0
		End If

		If BitPos < 8 + BitLen + NeededBits Then
			NeededBytes += 1     'Need an extra byte for bit sequence (closing match tag)
		End If

		If BitCt >= ByteCt + ByteLen + NeededBytes Then
			CheckSpace = True
		Else
			CheckSpace = False
		End If

	End Function

	Private Sub AddLiteralBits()

		If LitCnt = -1 Then Exit Sub

		AddBit(1)           'Then update bitstream with Literal Tag
		AddBit(LitCnt, False) 'And Literal count

	End Sub

	Private Sub AddBit(Bit As Integer, Optional Tag As Boolean = True)

		Bits = Buffer(BitCt) * 256  'Load last bitstream byte to upper byte of Bits

		If Tag = True Then  '1 bit to be added
			Bits += Bit * 2 ^ BitPos
			BitPos -= 1
		Else
			Bit = Bit * (2 ^ (8 - MaxBit))   '000xxxxx -> xxxxx000
			For I = 0 To MaxBit - 1  '5 bits to be added
				Bit = Bit * 2       'xxxxx000 -> x-xxxx0000
				Bits += Int(Bit / 256) * 2 ^ BitPos
				BitPos -= 1
				Bit = Bit Mod 256
			Next
		End If

		Buffer(BitCt) = Int(Bits / 256) 'Save upper byte to buffer

		If BitPos < 8 Then  'Lower byte is affected, save it to buffer, too
			BitCt -= 1
			Buffer(BitCt) = Bits Mod 256
			BitPos += 8
		End If

	End Sub

	Private Sub AddDistantMatch()

		'MsgBox("Distant Match length:" + vbTab + MaxLen.ToString +
		'vbNewLine + "Distant Match address:" + vbTab + Hex(MaxOffset).ToString)

		If (LitCnt = MaxLit) Or (LitCnt = -1) Then AddBit(0)   'add match tag if needed

		Buffer(ByteCt) = (MaxLen * 4) Or &H80           'Length of match*4+ Distant flag
		Buffer(ByteCt + 1) = MaxOffset Mod 256          'AdLo
		Buffer(ByteCt + 1) = Int(MaxOffset / 256)       'AdHi
		ByteCt += 3
		DtPos += MaxLen
		LitCnt = -1

	End Sub

	Private Sub AddFarMatch()

		'MsgBox("FarMatch Length: " + MaxLen.ToString +
		'vbNewLine + "FarMatch Offset: " + (MaxOffset - 1).ToString +
		'vbNewLine + "Buffer: " + (MaxLen * 4).ToString +
		'vbNewLine + "Buffer+1: " + ((MaxOffset - 1) Xor 255).ToString)

		If (LitCnt = MaxLit) Or (LitCnt = -1) Then AddBit(0)   'add match tag if needed

		Buffer(ByteCt) = MaxLen * 4                     'Length of match*4
		Buffer(ByteCt + 1) = (MaxOffset - 1) Xor &HFF     'To skip eor #$ff in code!
		ByteCt += 2
		DtPos += MaxLen
		LitCnt = -1

	End Sub

	Private Sub AddByte()
		'Update bitstream
		If LitCnt < MaxLit Then 'If LitCnt < Max
			LitCnt += 1         'Then increase it
		Else                    'If LitCnt = Max
			AddLiteralBits()
			LitCnt = 0          'This is the first byte of the new literal stream
		End If

		Buffer(ByteCt) = Prg(DtPos)
		ByteCt += 1

	End Sub

	Private Sub AddNearMatch()

		'MsgBox("NearMatch Length: " + (MaxLen).ToString + vbNewLine + "NearMatch Offset: " + (MaxOffset - 1).ToString)

		If (LitCnt = MaxLit) Or (LitCnt = -1) Then AddBit(0)   'add match tag if needed

		Buffer(ByteCt) = ((MaxOffset - 1) * 4) + MaxLen
		ByteCt += 1
		DtPos += MaxLen
		LitCnt = -1

	End Sub

	Private Sub CloseBuffer()

		If (LitCnt = MaxLit) Or (LitCnt = -1) Then AddBit(0)   'add match tag if needed
		'Close byte = &H0 - NOT needed (default value in empty buffer)

		If DtPos - MatchStart < LastByte - 2 Then
			If MatchStart + LastByte - 3 <= Prg.Length Then
				For I As Integer = 0 To LastByte - 3
					Buffer(2 + I) = Prg(MatchStart + I)
				Next
				Buffer(LastByte) = (Buffer(LastByte) Or &H40) And &HC0
				DtPos = MatchStart + LastByte - 2
				BitCt = LastByte - 2
			End If
		End If

		If (BlockCnt = 0) And (BufferCt = 0) Then       'If this is the first block in a file
			Distant = DtPos - 1                         'Save Data Pointer for distant match check
			'MsgBox("Distant Buffer lenght:" + vbTab + Hex(Distant).ToString)
		End If

		BufferCt += 1

		UpdateByteStream()

		ReDim Buffer(255)

		Buffer(0) = (BuffAdd + DtPos - 2) Mod 256
		Buffer(1) = (Int((BuffAdd + DtPos - 2) / 256)) Mod 256

		ByteCt = 2
		BitCt = 255
		LastByte = BitCt

		BitPos = 15
		AddBits(IOBit, 1)   'IO bit
		AddBits(0, 1)       'Compression bit

		LitCnt = -1
		MatchStart = DtPos

	End Sub

	Public Sub UpdateByteStream()   'THIS IS ALSO USED BY LZ4+RLE!!!

		ReDim Preserve ByteSt(BufferCt * 256 - 1)

		For I = 0 To 255
			ByteSt((BufferCt - 1) * 256 + I) = Buffer(I)
		Next

	End Sub

End Module
