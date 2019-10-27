Friend Module ModLZ
    Private CMStart, POffset, MLen As Integer
    Private ReadOnly MaxLongLen As Integer = 254    'Cannot be 255, there is an INY in the decompression ASM code, and that would make YR=#$00
    Private ReadOnly MaxMidLen As Integer = 61      'Cannot be more than 61 because 62=LongMatchTag, 63=NextFileTage
    Private ReadOnly MaxShortLen As Integer = 3     '1-3, cannot be 0 because it is preserved for EndTag
    Private ReadOnly DoDebug As Boolean = False
    Private ReadOnly MaxLit As Integer = 1 + 4 + 8 + 32 - 1  '=44

    Private LastByteCt As Integer = 255
    Private LastMOffset As Integer = 0
    Private LastMLen As Integer = 0

    Private ReadOnly LongMatchTag As Byte = &HF8    'Could be changed to &H00, but this is more economical
    Private ReadOnly NextFileTag As Byte = &HFC
    Private ReadOnly EndTag As Byte = 0             'Could be changed to &HF8, but this is more economical (Number of EndTags > Number of LongMatchTags)

    Private FirstOfNext As Boolean = False          'If true, this is the first block of next file in same buffer, Lit Selector Bit NOT NEEEDED
    Private NextFile As Boolean = False             'Indicates whether the next file is added to the same buffer

    Private BlockUnderIO As Integer = 0
    Private AdLoPos As Byte, AdHiPos As Byte
    Public Sub LZ(PN As Object, Optional FA As String = "", Optional FO As String = "", Optional FL As String = "", Optional FUIO As Boolean = False)
        On Error GoTo Err

        Dim ByC, BiC As Integer
        Dim FAN, FON, FLN As Integer

        If TypeOf PN Is Byte() Then
            'This does not need to be trimmed!!!
            Prg = PN
            FileUnderIO = FUIO
            PrgAdd = ConvertHexStringToNumber(FA)
            PrgLen = Prg.Length     'ConvertHexStringToNumber(FL)

            'MsgBox(Hex(PrgAdd) + vbNewLine + Hex(PrgLen))

        ElseIf TypeOf PN Is String Then
            ReDim Prg(0)

            'Read file to prg()
            If Strings.Right(PN, 1) = "*" Then
                If IO.File.Exists(Left(PN, Len(PN) - 1)) = True Then
                    Prg = IO.File.ReadAllBytes(Left(PN, Len(PN) - 1))
                    FileUnderIO = True
                Else
                    MsgBox(PN + vbNewLine + vbNewLine + "does not exist!", vbInformation + vbOKOnly)
                    Exit Sub
                End If
            Else
                If IO.File.Exists(PN) = True Then
                    Prg = IO.File.ReadAllBytes(PN)
                    FileUnderIO = False
                Else
                    MsgBox(PN + vbNewLine + vbNewLine + "does not exist!", vbInformation + vbOKOnly)
                    Exit Sub
                End If
            End If

            'Update File Start Address
            If FA = "" Then
                FAN = Prg(1) * 256 + Prg(0)
            Else
                FAN = ConvertHexStringToNumber(FA)
            End If
            PrgAdd = FAN

            'Update File Start Offset
            If FO = "" Then
                FON = 2
            Else
                FON = ConvertHexStringToNumber(FO)
            End If

            'Update File Length
            If FL = "" Then
                FLN = Prg.Length - FON
            Else
                FLN = ConvertHexStringToNumber(FL)
                If FLN + FON > Prg.Length Then
                    FLN = Prg.Length - FLN
                End If
            End If
            PrgLen = FLN    'Prg.Length


            'Trim prg from offset to a length of FLN
            For I As Integer = 0 To FLN - 1
                Prg(I) = Prg(FON + I)
            Next
            ReDim Preserve Prg(FLN - 1)
        End If

        LastByte = ByteCt   'ByteCt points at the next empty byte in buffer

NextBuffer:
        BlockUnderIO = 0       'Reset BlockUnderIO Flag

        If ((BlockCnt = 0) And (ByteCt = 254)) Or (ByteCt = 255) Then

            'First Block, ByteCt starts at 254, all other blocks: 255

            'Buffer is empty: we do not need ByC and BiC

            FirstOfNext = False                                         'First block in buffer, Lit Selector Bit is needed(?)
            NextFile = False                                            'This is the first file that is being added to an empty buffer
FileAddress:

            Buffer(ByteCt) = (PrgAdd + PrgLen - 1) Mod 256              'Add Address Info
            AdLoPos = ByteCt

            If (FileUnderIO = True) And (CheckIO(PrgLen - 1) = 1) Then  'Check if last byte of block is under IO
                BlockUnderIO = 1                                        'Yes, set BUIO flag
                ByteCt -= 1                                             'And skip 1 byte (=0) for IO Flag
            End If

            Buffer(ByteCt - 1) = Int((PrgAdd + PrgLen - 1) / 256)
            AdHiPos = ByteCt - 1

            ByteCt -= 2
            LitCnt = -1                                                 'Reset LitCnt here

        Else
            'Buffer is not empty, we need New File Match Tag, check if we have enough space left in open buffer

            '4 bytes and 0-1 bits needed for NextFileTag, Address Bytes and first Lit byte (+1 more if UIO)

            ByC = 4                                                     'ByC=5 only if BlockUnderIO=true - checked at DataFits()
            IIf((LitCnt = -1) Or (LitCnt = MaxLit), BiC = 1, BiC = 0)   'Calculate number of bits needed for Match Tag for new file

            If DataFits(ByC, BiC, LitCnt, CheckIO(PrgLen - 1)) Then

                'Buffer has enough space for New File Match Tag and New File Info and first Literal byte (and IO flag if needed)

                If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)  'Add Match Selector Bit only if needed
                Buffer(ByteCt) = NextFileTag                               'Then add New File Match Tag
                ByteCt -= 1
                FirstOfNext = True                                      'First block of next file in same buffer, Lit Selector Bit NOT NEEEDED
                NextFile = True                                         'Next file is being added to buffer that already has data
                GoTo FileAddress                                        'Then add Address Info
            Else
                'Next File Info does not fit, so close buffer
                CloseBuff()                                             'Then close and reset Buffer
                GoTo NextBuffer                                         'Then add Address Info to new buffer
            End If
        End If

        LastByte = ByteCt                       'The first byte of the ByteStream after (BlockCnt and IO Flag and) Address Bytes (251..253)

        'BufferCt = 0                           'First buffer of new file
        MatchStart = PrgLen - 1    '2
        Buffer(ByteCt) = Prg(PrgLen - 1)        'Add last byte of Prg to buffer
        LitCnt += 1
        ByteCt -= 1

        For POffset = PrgLen - 2 To 0 Step -1   'Skip Last byte

Restart:
            If FindMatch() = True Then

                'MATCH
                'Longest Match found, now check if it is a Longmatch, a Farmatch or a ShortMatch
                If MaxLen > MaxMidLen Then
                    'MaxLen=63-255
                    If LongMatch() = False Then GoTo Restart        'FALSE means nothing fit in buffer, new buffer was started
                ElseIf MaxLen > MaxShortLen Then
                    'MaxLen=4-62
                    If MidMatch() = False Then GoTo Restart
                Else
                    'MaxLen=1-3
                    If MaxOffset > 64 Then
                        'MaxOffset>64
                        If MidMatch() = False Then GoTo Restart
                    Else
                        'MaxOffset=<64
                        If ShortMatch() = False Then GoTo Restart
                    End If
                End If
            Else
                If Literal() = False Then GoTo Restart
            End If
        Next

        'Done with file, check literal counter, and update bitstream if LitCnt > -1
        If LitCnt > -1 Then AddLitBits()

        'All Lit Bytes and Bits are now added, but LitCnt is NOT reset here (needed for next match tag in buffer if another file is added)

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function FindMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("FindMatch")

        'CMStart = If(POffset + 256 > MatchStart, MatchStart, POffset + 256)     '255? vs 256?

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
                If Prg(POffset - MLen) = Prg(CmPos - MLen) And (MLen <= MaxLongLen) = True Then GoTo NextByte
            End If

            'Calculate length of matches
            If MLen > 1 Then
                MLen -= 1
                If (CmPos - POffset > 64) And (MLen < 2) Then
                    'Select Case LitCnt
                    'Case 0, 3, 4, 11, 12, 43, 44                         'RangeMax, RangeMax-1 - would go to next range
                    'GoTo AddMatch
                    'End Select
                    'MidMatch, but too short (2 bytes only)
                    'Ignore it for now, will be rechecked at CheckShortMatch
                Else
AddMatch:
                    'Save it to Match List
                    MatchCnt += 1
                    ReDim Preserve MatchOffset(MatchCnt)
                    ReDim Preserve MatchLen(MatchCnt)
                    ReDim Preserve MatchSave(MatchCnt)
                    MatchOffset(MatchCnt) = CmPos - POffset 'Save offset
                    MatchLen(MatchCnt) = MLen               'Save len

                    If MLen > MaxMidLen Then                       'Calculate and save MatchSave
                        'LongMatch (MLen=62-254)
                        MatchSave(MatchCnt) = MLen - 2
                    ElseIf MLen > MaxShortLen Then
                        'MidMatch (MLen=04-61)
                        MatchSave(MatchCnt) = MLen - 1
                    Else    'MLen=01-03
                        If MatchOffset(MatchCnt) > 64 Then
                            'MidMatch
                            MatchSave(MatchCnt) = MLen - 1
                        Else
                            'ShortMatch
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

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function LongMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("LongMatch")
        Dim Bits As Integer
        LongMatch = True
        'LONGMATCH

        If LitCnt > -1 Then CheckShortMidMatch()

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

            If DataFits(3, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                'Add long match here
                AddLitBits()    'First close last literal sequence
                AddLongM()      'Then add Long Match
            Else    'Long match does not fit, check if mid match would fit
                If MaxLen > MaxMidLen Then MaxLen = MaxMidLen           'Adjust MaxLen for MidMatch
                If DataFits(2, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                    'Add mid match here
                    AddLitBits()
                    AddMidM()
                    GoTo BufferFull2
                Else
                    If MaxLen > MaxShortLen Then MaxLen = MaxShortLen     'Adjust MaxLen for ShortMatch
                    If (MaxOffset <= 64) And (DataFits(1, Bits, LitCnt, CheckIO(POffset - MaxLen + 1))) Then
                        'Add short match here
                        AddLitBits()
                        AddShortM()
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
                        If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) Then
                            AddLitByte()        'Add Byte As literal And update LitCnt
                            GoTo BufferFull     'Then close buffer
                        Else
                            'Nothing fits, close buffer
BufferFull:
                            AddLitBits()
BufferFull2:
                            CloseBuff()
                            LongMatch = False    'Signal BufferFull and Goto Restart
                        End If
                    End If
                End If
            End If
        Else    'No literals, we need match tag (or literal tag)
            If DataFits(3, 1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                AddLongM()
            Else
                If MaxLen > MaxMidLen Then MaxLen = MaxMidLen
                If DataFits(2, 1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                    'no, check if first 63 bytes fit a far match
                    AddMidM()       'Then add Far Bytes
                    GoTo BufferFull2
                Else
                    If MaxLen > MaxShortLen Then MaxLen = MaxShortLen
                    If (MaxOffset <= 64) And DataFits(1, 1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                        'No, check if offset<64 and if first 3 bytes fit a near match
                        AddShortM()
                        GoTo BufferFull2
                        'IF ShortMatch DOES NOT FIT - LITBYTE WILL NOT FIT
                        'LITBYTE WOULD NEED AT LEAST 2 BITS And 1 Byte, ShortMatch NEEDS 1 BIT And 1 BYTE
                    Else
                        'Nothing fits, close buffer (no literals)
                        GoTo BufferFull2
                    End If
                End If
            End If
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function MidMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("MidMatch")
        Dim Bits As Integer
        MidMatch = True
        'MIDMATCH

        If LitCnt > -1 Then CheckShortMidMatch()

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
            If DataFits(2, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                AddLitBits()        'First close last literal sequence
                AddMidM()           'Then add Far Bytes
            Else
                If MaxLen > MaxShortLen Then MaxLen = MaxShortLen
                If (MaxOffset <= 64) And DataFits(1, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                    AddLitBits()
                    AddShortM()
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
                    If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) Then
                        AddLitByte()        'Add Byte As literal And update LitCnt
                        GoTo BufferFull     'Then close buffer
                    Else
                        'Nothing fits, close buffer
BufferFull:
                        AddLitBits()
BufferFull2:
                        CloseBuff()
                        MidMatch = False    'Goto Restart
                    End If
                End If
            End If
        Else    'No literals, we need match tag (or literal tag)
            If DataFits(2, 1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                AddMidM()       'Then add Far Bytes
            Else
                If MaxLen > MaxShortLen Then MaxLen = MaxShortLen
                If (MaxOffset <= 64) And DataFits(1, 1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                    'Check if offset<64 and if first 3 bytes fit a near match
                    AddShortM()
                    GoTo BufferFull2
                Else
                    'Nothing fits, close buffer (no literals)
                    GoTo BufferFull2
                End If
            End If
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function ShortMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("ShortMatch")
        ShortMatch = True

        If LitCnt > -1 Then CheckShortMidMatch()

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
            'We have already reserved space for literal bits, check is match fits
            If DataFits(1, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                'Yes, we have space
                AddLitBits()    'First close last literal sequence
                AddShortM()       'Then add Near Byte
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
                If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) Then
                    'Match does not fit, check if we can add byte as literal
                    AddLitByte()           'Add byte as literal and update LitCnt
                    GoTo BufferFull     'Then close buffer
                Else
                    'Nothing fits, close buffer
BufferFull:
                    AddLitBits()
BufferFull2:
                    CloseBuff()
                    ShortMatch = False    'Goto Restart
                End If
            End If
        Else
            If DataFits(1, 1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                'Yes, we have space
                AddShortM()       'Then add Near Byte
            Else
                'If does not fit as Near Match, do not attempt to add as literal, as this would be the first literal, and would need more space
                'then as a match (1 byte + 1 bits as match) vs (1 byte + 2 bits as literal)
                'Nothing fits, close buffer (no literals)
                GoTo BufferFull2
            End If
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function Literal() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("Literal")

        Literal = True
        'LITERAL

        'CheckShortMidMatch()      'THIS WILL NOT IMPROVE COMPRESSION

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

        If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) = True Then 'We are adding 1 literal byte and 1+5 bits
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

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        Literal = False

    End Function

    Private Function DataFits(ByteLen As Integer, BitLen As Integer, Literals As Integer, Optional SequenceUnderIO As Integer = 0) As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("DataFits")
        'Do not include closing bits in function call!!!

        Dim NeededBytes As Integer = 1  'Close Byte
        Dim CloseBit As Integer = 0     'CloseBit length - Match selector bit length, not needed most of the time, cannot overlap with Close Byte!!!

        'If FileUnderIO = True Then                              'Some blocks of the file will need to go under IO
        'NeededBytes += BlockUnderIO Or SequenceUnderIO      'If this block does or would go under IO then add 1
        'End If

        If (FileUnderIO = True) And (BlockUnderIO = 0) And (SequenceUnderIO = 1) Then

            'Some blocks of the file will go UIO, but sofar this block is not UIO, and next byte is the first one that goes UIO

            NeededBytes += 1    'Need an extra byte if the next byte in sequence is the first one that goes UIO
        End If

        'Check if we have pending Literals
        'If no pending Literals, or Literals=MaxLit, then we need to save 1 bit for match tag
        'Otherwise, next item must be a match, we do not need a match tag
        If (Literals = -1) Or (Literals = MaxLit) Then CloseBit = 1

        If BitLen >= 8 Then     'Can be up to 11 bits
            NeededBytes += 1    'Need 1 more byte
            BitLen -= 8
        End If

        If BitPos < 8 + BitLen + CloseBit Then
            NeededBytes += 1     'Need an extra byte for bit sequence (closing match tag)
        End If

        If ByteCt >= BitCt + ByteLen + NeededBytes Then  '>= because NeededBytes also includes Close Byte
            DataFits = True
            'Data will fit
            If (FileUnderIO = True) And (BlockUnderIO = 0) And (SequenceUnderIO = 1) And (LastByte <> ByteCt) Then

                'This is the first byte in the block that will go UIO, so lets update the buffer to include the IO flag

                For I As Integer = ByteCt To AdHiPos            'Move all data to the left in buffer, including AdHi
                    Buffer(I - 1) = Buffer(I)
                Next
                Buffer(AdHiPos) = 0                             'IO Flag to previous AdHi Position
                ByteCt -= 1                                     'Update ByteCt to next empty position in buffer
                AdHiPos -= 1                                    'Update AdHi Position in Buffer
                BlockUnderIO = 1                                'Set BlockUnderIO Flag
            End If
        Else
            DataFits = False
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        DataFits = False

    End Function

    Private Sub AddLongM()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddLongM")

        If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCt) = LongMatchTag                   'Long Match Flag = &HF8
        Buffer(ByteCt - 1) = MaxLen
        Buffer(ByteCt - 2) = MaxOffset - 1
        ByteCt -= 3

        POffset -= MaxLen

        ResetLit()

        LM += 1

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddMidM()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddMidM")

        If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCt) = MaxLen * 4                         'Length of match (#$02-#$3f, cannot be #$00 (end byte), and #$01 - distant selector??)
        Buffer(ByteCt - 1) = MaxOffset - 1
        ByteCt -= 2

        POffset -= MaxLen

        ResetLit()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddShortM()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddShortM")

        If (LitCnt = -1) Or (LitCnt = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCt) = ((MaxOffset - 1) * 4) + MaxLen
        ByteCt -= 1

        POffset -= MaxLen

        ResetLit()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddLitByte()
        On Error GoTo Err

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

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub CheckShortMidMatch()
        On Error GoTo Err

        Dim MO, ML As Integer
        'Finds 2-byte MidMatches where it would save on bit length

        Select Case LitCnt
            Case 0
                If OverMaxLit = True Then       'LitCnt=45 actually
                    'OverMax:
                    For I As Integer = POffset + 2 To MatchStart - 1
                        If (Prg(POffset + 1) = Prg(I)) And (Prg(POffset + 2) = Prg(I + 1)) And (I < POffset + 1 + 256) Then
                            MO = MaxOffset
                            ML = MaxLen
                            MaxOffset = I - (POffset + 1)
                            MaxLen = 1
                            LitCnt = LitCnt + MaxLit + 1 - 2 '= 0/1 + 44 + 1 - 2 = 43
                            ByteCt += 2             'Step back 2 bytes
                            POffset += MaxLen + 1   'Repositon POffset

                            Buffer(BitCt) = 0       'Delete unused bit stream data (deleting 8+1 bits)
                            BitCt -= 1              'Step back 8+1 bits
                            BitPos += 1             'we have already added 9 bits and 45 (0-44) Literal bytes to the bit and byte streams
                            If BitPos > 15 Then     'So we are now removing 8+1 bits from the bit stream, and also update the BitPos pointer
                                BitPos -= 8
                                Buffer(BitCt) = 0   'Delete unused bit steam data
                                BitCt -= 1
                            End If
                            Buffer(BitCt) = Buffer(BitCt) And (&H100 - (2 ^ (BitPos - 7)))  'Delete bit stream
                            AddLitBits()            'We are re-adding bits for 43 Literal bytes
                            AddMidM()               'This will also reset LitCnt and OverMaxLit
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
                'Case 1
                'If OverMaxLit = True Then
                'GoTo OverMax
                'Else
                'GoTo UnderMax
                'End If
            Case 1, 2, 5, 6, 13, 14                         'rangemax+1,rangemax+2 (value cannot be 0!!!) (rangemax: 0,4,12,44)
                'UnderMax:
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
                        AddMidM()
                        SM2 += 1
                        MaxOffset = MO
                        MaxLen = ML
                        POffset -= 1
                        Exit Sub
                    End If
                Next
        End Select

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub CheckOverlap()
        On Error GoTo Err

        Dim PO As Integer = POffset + 1
        Dim MO, ML As Integer
        'WE HAVE 1 LITERAL AFTER A MATCH THAT IS LONGER THAN 2 BYTES, SO CHECK IF THERE IS AN OVERLAP BACKWARDS
        'CheckOverlap = True

        For I As Integer = PO + 1 To MatchStart - 1
            If I < PO + 256 Then
                'If current byte= byte(I) and previous byte (last byte of last match)=byte(I+1) then we have an overlap
                If (Prg(PO) = Prg(I)) And (Prg(PO + 1) = Prg(I + 1)) Then
                    LastMLen -= 1   'Last Match length will be 1 less as 
                    'Update Last Match
                    If I < PO + 65 Then                        'Overlap is a shot match
                        If DataFits(1, 1, -1, CheckIO(POffset)) Then  'Check if new 2-byte match will fit as a short match
                            If (LastMOffset <= 64) And (LastMLen <= MaxShortLen) Then   'Update previous match as a near match
                                Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
                                ByteCt = LastByteCt - 1
                            ElseIf LastMLen <= MaxMidLen Then                           'Update previous match as a far match
                                Buffer(LastByteCt) = LastMLen * 4
                                Buffer(LastByteCt - 1) = LastMOffset - 1
                                ByteCt = LastByteCt - 2
                            Else                                                        'Update previous match as a long match
                                Buffer(LastByteCt) = LongMatchTag
                                Buffer(LastByteCt - 1) = LastMLen
                                Buffer(LastByteCt - 2) = LastMOffset - 1
                                ByteCt = LastByteCt - 3
                            End If
                            MO = MaxOffset      'Save next match's offset
                            ML = MaxLen         'and length
                            MaxOffset = I - PO
                            MaxLen = 1
                            LitCnt = -1
                            AddShortM()
                            MaxOffset = MO
                            MaxLen = ML
                            POffset = PO - 1
                        End If
                    Else                                            'Overlap is a long match - THIS WOULD ONLY SAVE 1 BIT
                        If DataFits(2, 1, -1, CheckIO(POffset)) Then
                            If (LastMOffset <= 64) And (LastMLen <= MaxShortLen) Then            'Make it a near match
                                Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
                                ByteCt = LastByteCt - 1
                            ElseIf LastMLen <= MaxMidLen Then                                   'Make it a far match
                                Buffer(LastByteCt) = LastMLen * 4
                                Buffer(LastByteCt - 1) = LastMOffset - 1
                                ByteCt = LastByteCt - 2
                            Else                                                                'Make it a long match
                                Buffer(LastByteCt) = LongMatchTag
                                Buffer(LastByteCt - 1) = LastMLen
                                Buffer(LastByteCt - 2) = LastMOffset - 1
                                ByteCt = LastByteCt - 3
                            End If
                            MO = MaxOffset
                            ML = MaxLen
                            MaxOffset = I - PO
                            MaxLen = 1
                            LitCnt = -1
                            AddMidM()
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

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddLitBits()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddLitBits")

        If LitCnt = -1 Then Exit Sub

        If FirstOfNext = False Then AddRBits(1, 1)               'Add Literal Selector

        FirstOfNext = False

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

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddRBits(Bit As Integer, BCnt As Byte)
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddRBits:" + vbTab + Hex(Bit).ToString + vbTab + BCnt.ToString)

        Bits = Buffer(BitCt) * 256  'Load last bitstream byte to upper byte of Bits

        Bit *= 2 ^ (8 - BCnt)       'Shift bits to the leftmost position in byte

        'Add bits to bitstream
        For I = 1 To BCnt  '1-5 bits to be added
            Bit *= 2  'xxxxxx00 -> x-xxxxx000
            Bits += Int(Bit / 256) * (2 ^ BitPos)
            BitPos -= 1
            Bit = Bit Mod 256
        Next

        Buffer(BitCt) = Int(Bits / 256) 'Save upper byte to buffer

        If BitPos < 8 Then  'Lower byte is affected, save it to buffer, too
            BitCt += 1      'Next byte in BitStream
            Buffer(BitCt) = Bits Mod 256
            BitPos += 8     'Reset bitpos counter
        End If

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub CloseBuff()
        On Error GoTo Err

        If DoDebug Then Debug.Print("CloseBuff")

        Buffer(ByteCt) = EndTag

        'Close Byte Match Tag = 0, so it is not needed as it is the default value in the next bit position!!!

        Buffer(0) = Buffer(0) And &H7F                    'Delete Compression Bit (Default (i.e. compressed) value is 0)

        'FIND UNCOMPRESSIBLE BLOCKS (only applies to first file in buffer in which case LastByte points at the very first data byte in stream)

        If (MatchStart - POffset <= LastByte) And (MatchStart > LastByte - 1) And (NextFile = False) Then

            'Less than 252/253 bytes        AND  not the end of File     AND       No other files in this buffer
            LastByte = AdLoPos - 2

            'Check uncompressed Block IO Status
            If (CheckIO(MatchStart) Or CheckIO(MatchStart - (LastByte - 1)) = 1) And (FileUnderIO = True) Then
                'If the block will be UIO than only (Lastbyte-1) bytes will fit,
                'So we only need to check that many bytes
                Buffer(AdLoPos - 1) = 0 'Set IO Flag
                AdHiPos = AdLoPos - 2   'Updae AdHiPos
                LastByte = AdHiPos - 1  'Update LastByte
                'POffset += 1
            ElseIf (CheckIO(MatchStart - LastByte) = 1) And (FileUnderIO = True) Then
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

            POffset = MatchStart - LastByte                         'Update POffset

            Buffer(AdHiPos) = Int((PrgAdd + POffset) / 256)
            Buffer(AdLoPos) = (PrgAdd + POffset) Mod 256

            For I As Integer = 0 To LastByte - 1            '-1 because the first byte of the buffer is the bitstream
                Buffer(LastByte - I) = Prg(MatchStart - I)
            Next

            Buffer(0) = &H80                                        'Set Copression Bit to 1 (=Uncompressed block)
            ByteCt = 1

        End If

        BlockCnt += 1
        UpdateByteStream()

        ResetBuffer()               'Resets buffer variables

        NextFile = False            'Reset Next File flag

        If POffset < 0 Then Exit Sub 'We have reached the end of the file -> exit

        'If we have not reached the end of the file, then update buffer

        BlockUnderIO = CheckIO(POffset)          'Check if last byte of prg could go under IO

        Buffer(ByteCt) = (PrgAdd + POffset) Mod 256
        AdLoPos = ByteCt
        If (FileUnderIO = True) And (BlockUnderIO = 1) Then
            ByteCt -= 1
        End If
        Buffer(ByteCt - 1) = Int((PrgAdd + POffset) / 256) Mod 256
        AdHiPos = ByteCt - 1
        ByteCt -= 2
        LastByte = ByteCt               'LastByte = the first byte of the ByteStream after and Address Bytes (253 or 252 with blockCnt)

        Buffer(ByteCt) = Prg(POffset)   'Copy First Lit Byte to Buffer
        ByteCt -= 1                     'Update Byte Pos Counter
        MatchStart = POffset            'Update Match Start Flag
        POffset -= 1                    'Next Byte in Prg
        If POffset < 0 Then POffset = 0 'Unless it is <0		NEEDED DO NOT DELETE!!!
        LitCnt = 0                      'Literal counter has 1 value

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub CloseLastBuff()
        On Error GoTo Err

        If DoDebug Then Debug.Print("CloseLastBuff")

        Buffer(ByteCt) = EndTag

        'Close Byte Match Tag = 0, so it is not needed as it is the default value in the next bit position!!!

        Buffer(0) = Buffer(0) And &H7F                    'Delete Compression Bit (Default (i.e. compressed) value is 0)

        '--------------------------------------------------------------------------------------------------

        '--------------------------------------------------------------------------------------------------
        If NextFile = False Then
            'Check uncompressed Block IO Status
            If (CheckIO(MatchStart) Or CheckIO(0) = 1) And (FileUnderIO = True) Then
                'Block will go UIO
                'IO Flag will be set
                AdHiPos = AdLoPos - 2   'Update AdHiPos to include IO Flag
                LastByte = AdHiPos - 1  'Update LastByte
            Else
                'Block will not go UIO
                'IO Flag will not be set
                AdHiPos = AdLoPos - 1   'Update AdHiPos, IO Flag is not needed
                LastByte = AdHiPos - 1  'UpdateLastByte
            End If

            'Calculate available space and number of bytes left...

            Dim BytesForBitStream As Integer = 1                                '1 byte needed for bitStream if block is left uncompressed
            Dim BytesLeft As Integer = MatchStart + 1                           '0...MatchStart
            Dim BytesAvailable As Integer = LastByte + 1 - BytesForBitStream    '1...LastByte

            If BytesLeft <= BytesAvailable Then
                'If we have enough space for the uncompressed version of the block then
                'store block uncompressed
                'Otherwise, do not touch it

                ReDim Buffer(255)

                Buffer(AdHiPos) = Int((PrgAdd - 1) / 256)           'New forward block address Hi
                Buffer(AdLoPos) = (PrgAdd - 1) Mod 256              'New forward block address Lo

                For I As Integer = 0 To MatchStart                  'Copy remaining bytes from file
                    Buffer(BytesForBitStream + I) = Prg(I)          'leave 1 byte for bitstream
                Next

                If BytesLeft < BytesAvailable Then
                    Buffer(LastByte) = BytesLeft                    'If block is not full then save number of bytes+1
                    Buffer(0) = &HC0                                'Set Length Flag and Compression Bit to 1 (=Uncompressed Short Block)
                Else
                    Buffer(0) = &H80                                'Set Copression Bit to 1 (=Uncompressed block)
                End If

            End If
        End If

        BlockCnt += 1
        UpdateByteStream()

        ResetBuffer()               'Resets buffer variables

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub UpdateByteStream()   'THIS IS ALSO USED BY LZ4+RLE!!!
        On Error GoTo Err

        ReDim Preserve ByteSt(BlockCnt * 256 - 1)

        For I = 0 To 255
            ByteSt((BlockCnt - 1) * 256 + I) = Buffer(I)
        Next

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub ResetLit()
        On Error GoTo Err

        LitCnt = -1
        OverMaxLit = False

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SaveLastMatch()
        On Error GoTo Err

        LastByteCt = ByteCt
        LastMOffset = MaxOffset
        LastMLen = MaxLen

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub ResetBuffer()
        On Error GoTo Err

        ReDim Buffer(255)       'New empty buffer

        'Initialize variables

        If BlockCnt = 0 Then
            ByteCt = 254        'First file in part, last byte of first block will be BlockCnt
        Else
            ByteCt = 255        'All other blocks:	Bytestream start at the end of the buffer
        End If

        BitCt = 0               'Bitstream always starts at the beginning of the buffer

        BitPos = 15             'Reset Bit Position Counter

        'DO NOT RESET LitCnt HERE!!! It is needed for match tag check

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function CheckIO(Offset As Integer) As Integer
        On Error GoTo Err

        CheckIO = If((PrgAdd + Offset >= &HD000) And (PrgAdd + Offset <= &HDFFF) And (FileUnderIO = True), 1, 0)

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

End Module
