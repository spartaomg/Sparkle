﻿Friend Module ModFasterPacker

    Private CMStart, CMEnd, POffset, MLen As Integer
    Private ReadOnly MaxLongLen As Integer = 254    'Cannot be 255, there is an INY in the decompression ASM code, and that would make YR=#$00
    Private ReadOnly MaxMidLen As Integer = 61      'Cannot be more than 61 because 62=LongMatchTag, 63=NextFileTage
    Private ReadOnly MaxShortLen As Integer = 3     '1-3, cannot be 0 because it is preserved for EndTag
    Private ReadOnly DoDebug As Boolean = False
    Private ReadOnly MaxLit As Integer = 1 + 4 + 8 + 32 - 1  '=44 - this seems to be optimal, 1+4+8+16 and 1+4+8+64 are worse...

    Public ReadOnly LongMatchTag As Byte = &HF8    'Could be changed to &H00, but this is more economical
    Public ReadOnly NextFileTag As Byte = &HFC
    Public ReadOnly EndTag As Byte = 0             'Could be changed to &HF8, but this is more economical (Number of EndTags > Number of LongMatchTags)

    Private FirstBlockOfNextFile As Boolean = False 'If true, this is the first block of next file in same buffer, Lit Selector Bit NOT NEEEDED
    Private NextFileInBuffer As Boolean = False     'Indicates whether the next file is added to the same buffer

    Private BlockUnderIO As Integer = 0
    Private AdLoPos As Byte, AdHiPos As Byte

    Private PreM As Boolean = False
    Private PostM As Boolean = False
    Private PreOL As Boolean = False
    Private PostOL As Boolean = False

    Public Sub NewLZ(PN As Byte(), Optional FA As String = "", Optional FUIO As Boolean = False)
        On Error GoTo Err
        'The only two parameters that are needed are FA and FUIO

        '----------------------------------------------------------------------------------------------------------
        'PROCESS FILE
        '----------------------------------------------------------------------------------------------------------

        Prg = PN
        FileUnderIO = FUIO
        PrgAdd = Convert.ToInt32(FA, 16)
        PrgLen = Prg.Length

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
        LitCnt = -1                                                 'Reset LitCnt here
        LastByte = ByteCnt                       'The first byte of the ByteStream after (BlockCnt and IO Flag and) Address Bytes (251..253)

        'BufferCt = 0                           'First buffer of new file
        MatchStart = PrgLen - 1    '2
        LastMS = MatchStart
        Buffer(ByteCnt) = Prg(PrgLen - 1)        'Add last byte of Prg to buffer
        LitCnt += 1
        ByteCnt -= 1
        LastPO = PrgLen - 2

        CompressFile(LastPO)

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub CompressFile(StartPos As Integer)
        On Error GoTo Err

        '----------------------------------------------------------------------------------------------------------
        'COMPRESS FILE
        '----------------------------------------------------------------------------------------------------------

        For POffset = StartPos To 0 Step -1   'Skip Last byte

Restart:

            If FindMatch() = True Then

                'MATCH
                'Longest Match found, now check if it is a Longmatch, a Midmatch or a ShortMatch
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
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function FindMatch(Optional PO As Integer = -1) As Boolean
        On Error GoTo Err
        'FINDS MATCHES AND IDENTIFIES THE ONE THAT SAVES THE MOST

        On Error GoTo Err

        If DoDebug Then Debug.Print("FindMatch")

        If PO = -1 Then PO = POffset

        'CMStart = If(POffset + 256 > MatchStart, MatchStart, POffset + 256)     '255? vs 256?

        CMStart = PO + 256

        If BlockCnt = 1 Then
            If CMStart > PrgLen - 1 Then CMStart = PrgLen - 1
        Else
            If CMStart > MatchStart Then CMStart = MatchStart
        End If

        ''If BlockCnt > 1 Then
        'If CMStart > MatchStart Then CMStart = MatchStart
        ''Else                                   'THIS DOES NOT SIGNIFICANTLY IMPROVE COMPRESSION BUT SLOWS DEPACKING DOWN
        ''If CMStart > PrgLen - 1 Then           'The first block is always loaded first, so the 2nd block can reach back to the 1st
        ''CMStart = PrgLen - 1                   'looking for matches
        ''End If                                 'Same applies to the last block which is always loaded last so it can search
        ''End If                                 'the next to last for matches

        '       PO                            CMPos (CMStart=PO+256 vs MatchStart)
        '    +++P|C...    <-      <-     <-+++C|
        '--------|-----------------------------|-------
        '       0|123456789A...            ...FF

        MatchCnt = 0
        ReDim MatchOffset(0), MatchLen(0), MatchSave(0), MatchType(0)
        MLen = -1

        For CmPos = CMStart To PO + 1 Step -1       'Offset
NextByte:
            MLen += 1                               'Match length, inital value set here to 0

            If PO - MLen > -1 Then
                If (Prg(PO - MLen) = Prg(CmPos - MLen)) And (MLen <= MaxLongLen) Then GoTo NextByte 'Found match, increase length
            End If

            'Calculate length of matches
            If MLen > 1 Then    'Match length of at least 2 bytes
                MLen -= 1
                If (CmPos - PO > 64) And (MLen = 1) Then
                    'Exclude 2 - Byte mid matches
                Else
                    'Save it to Match List
                    MatchCnt += 1
                    ReDim Preserve MatchOffset(MatchCnt)
                    ReDim Preserve MatchLen(MatchCnt)
                    ReDim Preserve MatchSave(MatchCnt)
                    ReDim Preserve MatchType(MatchCnt)

                    MatchOffset(MatchCnt) = CmPos - PO          'Save offset
                    MatchLen(MatchCnt) = MLen                   'Save len

                    If MLen > MaxMidLen Then                    'Calculate number of saved bytes and store it in MatchSave
                        'LongMatch (MLen=62-254) (actual length 63-255)
                        MatchSave(MatchCnt) = MLen - 2
                        MatchType(MatchCnt) = "l"
                    ElseIf MLen > MaxShortLen Then
                        'MidMatch (MLen=04-61) (actual length 05-62)
                        MatchSave(MatchCnt) = MLen - 1
                        MatchType(MatchCnt) = "m"
                    Else    'MLen=01-03 (actual length 02-04)
                        If MatchOffset(MatchCnt) > 64 Then
                            'MidMatch
                            MatchSave(MatchCnt) = MLen - 1
                            MatchType(MatchCnt) = "m"
                        Else
                            'ShortMatch
                            MatchSave(MatchCnt) = MLen - 0
                            MatchType(MatchCnt) = "s"
                        End If
                    End If
                End If
            End If
            MLen = -1   'Reset MLen
        Next

        If MatchCnt > 0 Then
            'MATCHES FOUND, IDENTIFY MOST ECONOMIC ONE

            'Reset MaxS variables for max short match
            MaxSLen = 0
            MaxSOff = 0
            'Find longest match in Match List
            MaxSave = MatchSave(1)
            MaxLen = MatchLen(1)
            MaxOffset = MatchOffset(1)
            MaxType = MatchType(1)
            Match = 1
            For I = 1 To MatchCnt
                If MatchSave(I) = MaxSave Then      'If next item saves the same number of bytes,

                    If MatchLen(I) > MaxLen Then    'Update, if match sequence is longer
                        MaxSave = MatchSave(I)
                        MaxLen = MatchLen(I)
                        MaxOffset = MatchOffset(I)
                        MaxType = MatchType(I)
                        Match = I
                    End If

                ElseIf MatchSave(I) > MaxSave Then  'Otherwise, update if saves more (decompression is faster this way)

                    MaxSave = MatchSave(I)
                    MaxLen = MatchLen(I)
                    MaxOffset = MatchOffset(I)
                    MaxType = MatchType(I)
                    Match = I

                End If
                'This will save the longest short match in case a longer match would not fit in the buffer
                If (MatchOffset(I) <= 64) And (MatchLen(I) < 4) And (MatchLen(I) > MaxSLen) Then
                    MaxSLen = MatchLen(I)
                    MaxSOff = MatchOffset(I)
                End If
            Next

            If MaxSave = 0 Then
                MaxLen = 0
                FindMatch = False
            Else
                FindMatch = True
            End If
        Else
            MaxLen = 0
            FindMatch = False
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function ReplaceMidWithShort() As Boolean
        On Error GoTo Err

        Dim Before As String = ""
        Dim After As String = ""

        'For I As Integer = 0 To LastMLen
        'Before = Hex(Prg(LastPOffset + LastMOffset - I)) + " " + Before
        'Next
        'Before = "| " + Before
        'For I As Integer = 0 To MaxLen
        'Before = Hex(Prg(POffset + MaxOffset - I)) + " " + Before
        'Next

        ReplaceMidWithShort = False

        Dim LastPO As Integer = LastPOffset
        Dim TmpPO As Integer = POffset - MaxLen

        Dim ReplaceLength As Byte = MaxLen + LastMLen + 2
        If ReplaceLength < 6 Then GoTo NoReplace

        Dim MaxL As Integer = 1         'Max match length found
        Dim MaxO As Integer = 0         'Offset of match found
        Dim MaxSL As Integer = If(ReplaceLength >= 8, 4, ReplaceLength - 4)        'Max length of short match we need

        Dim FrstML, FrstMO As Integer
        Dim ScndML, ScndMO As Integer
        Dim ThrdML, ThrdMO As Integer

        'Offset = 1 to 64 or what's left of prg.length
        For O As Integer = If(Prg.Length - 1 - LastPOffset < 64, Prg.Length - 1 - LastPOffset, 64) To 1 Step -1
            'Length = 0 to MaxSL or what's left till prg start
            For L As Integer = 0 To If(LastPOffset >= MaxSL, MaxSL, LastPOffset)
                If Prg(LastPOffset - L) <> Prg(LastPOffset + O - L) Then
                    If L > MaxL Then
                        'Match of at least 2 bytes
                        MaxL = L
                        MaxO = O
                    End If
                    Exit For
                ElseIf L = MaxSL Then
                    MaxL = L
                    MaxO = O
                    Exit For
                End If
            Next
            If MaxL = MaxSL Then
                'We have found a max length match, search is complete
                Exit For
            End If
        Next

        'No match, matches cannot be replaced
        If MaxL < 2 Then GoTo NoReplace

        'Update shearch variables
        ReplaceLength -= MaxL

        'Remaining length is too long for 2 short matches
        If ReplaceLength > 8 Then GoTo NoReplace

        'Save first short match parameters
        FrstML = MaxL
        FrstMO = MaxO

        'For L As Integer = 0 To FrstML - 1
        'After = Hex(Prg(LastPOffset + FrstMO - L)) + " " + After
        'Next
        'After = "| " + After

        MaxSL = If(ReplaceLength >= 6, 4, ReplaceLength - 2)

        If MaxSL < 2 Then GoTo NoReplace

        LastPOffset -= FrstML

        MaxL = 0
        MaxO = 0

        'Offset = 1 to 64 or what's left of prg.length
        For O As Integer = 1 To If(Prg.Length - 1 - LastPOffset < 64, Prg.Length - 1 - LastPOffset, 64)
            'Length = 0 to MaxSL or what's left till prg start
            For L As Integer = 0 To If(LastPOffset >= MaxSL, MaxSL, LastPOffset)
                If Prg(LastPOffset - L) <> Prg(LastPOffset + O - L) Then
                    If Prg(LastPOffset - L) <> Prg(LastPOffset + O - L) Then
                        If L > MaxL Then
                            'Match of at least 2 bytes
                            MaxL = L
                            MaxO = O
                        End If
                        Exit For
                    ElseIf L = MaxSL Then
                        MaxL = L
                        MaxO = O
                        Exit For
                    End If
                End If
            Next
            If MaxL = MaxSL Then
                'We have found a max length match, search is complete
                Exit For
            End If
        Next

        'No match, matches cannot be replaced
        If MaxL < 2 Then GoTo NoReplace

        ReplaceLength -= MaxL

        'More than 4 bytes left, longer than short match, replacement cannot be done
        If ReplaceLength > 4 Then GoTo NoReplace

        'Save second short match parameters
        ScndML = MaxL
        ScndMO = MaxO

        'For I As Integer = 0 To MaxL - 1
        'After = Hex(Prg(LastPOffset + MaxO - I)) + " " + After
        'Next
        'After = "| " + After

        LastPOffset -= ScndML

        MaxL = 0
        MaxO = 0

        'Offset = 1 to 64 or what's left of prg.length
        For O As Integer = 1 To If(Prg.Length - 1 - LastPOffset < 64, Prg.Length - 1 - LastPOffset, 64)
            'Length = 0 to 4 or what's left till prg start
            For L As Integer = 0 To If(LastPOffset >= ReplaceLength, ReplaceLength, LastPOffset)
                If Prg(LastPOffset - L) <> Prg(LastPOffset + O - L) Then
                    If L >= ReplaceLength Then
                        MaxL = ReplaceLength
                        MaxO = O
                    End If
                    Exit For
                End If
            Next
            If MaxL = ReplaceLength Then
                Exit For
            End If
        Next

        'Last Match must be = remaining length to be replaced
        If MaxL <> ReplaceLength Then GoTo NoReplace

        'Save third short match parameters
        ThrdML = MaxL
        ThrdMO = MaxO


        'For I As Integer = 0 To MaxL - 1
        'After = Hex(Prg(LastPOffset + MaxO - I)) + " " + After
        'Next

        'Debug.Print(Before + vbTab + After)

        'Debug.Print("Replacement found: " + (MaxLen + 1).ToString + "(" + MaxOffset.ToString + ") +" +
        '			(LastMLen + 1).ToString + "(" + LastMOffset.ToString + ") " + vbTab +
        'FrstML.ToString + "(" + FrstMO.ToString + ") " +
        '			"+" + ScndML.ToString + "(" + ScndMO.ToString + ") " +
        '			"+" + ThrdML.ToString + "(" + ThrdMO.ToString + ") ")

        'GoTo NoReplace
        '------------------------------------------------------------------
        '	WE HAVE ALL 3 SHORT MATCHES TO REPLACE TO MID MATCHES WITH
        '------------------------------------------------------------------

        FrstML -= 1
        ScndML -= 1
        ThrdML -= 1

        Buffer(ByteCnt + 2) = ((FrstMO - 1) * 4) + FrstML
        AddRBits(0, 1)
        Buffer(ByteCnt + 1) = ((ScndMO - 1) * 4) + ScndML
        AddRBits(0, 1)
        Buffer(ByteCnt + 0) = ((ThrdMO - 1) * 4) + ThrdML

        LastByteCt = ByteCnt
        LastMOffset = ThrdMO
        LastMLen = ThrdML
        LastMType = "s"

        ByteCnt -= 1

        POffset = LastPOffset - ThrdML

        'Debug.Print(TmpPO = POffset)

        ReplaceMidWithShort = True

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

NoReplace:
        'Restore LastPOffset
        LastPOffset = LastPO

    End Function

    Private Function Find2ByteMatches() As Boolean
        On Error GoTo Err

        Find2ByteMatches = False

        Select Case LitCnt Mod (MaxLit + 1)
            Case 0, 1, 2, 3, 4, 5, 6, 13, 14

                If LitCnt > 0 Then
                    PreM = PreMatch()        'Must be 2-byte MidMatch stored in 2 bytes
                    PostM = PostMatch()      'Must be 2-byte MidMatch stored in 2 bytes
                Else
                    PreM = False
                    PostM = False
                End If
                PreOL = PreOverlap()         'Must be 2-byte ShortMatch stored in 1 byte
                PostOL = PostOverlap()       'Must be 2-byte ShortMatch stored in 1 byte

                If (PreM = False) And (PostM = False) And (PreOL = False) And (PostOL = False) Then Exit Function

            Case Else
                Exit Function
        End Select

        'RangeMins:  0,1,5,13
        Select Case LitCnt Mod (MaxLit + 1)
            Case 0          '1B2b
                '0/5/13+0                                                                       Lits    Pre     Post    Saved
                If LitCnt = 0 Then
                    If PreOL = True Then                                '1      0/5/13 -> -1/4/12   5B5b    1B1b            1 bit
                        SavePreOverlap()
                        BitsSaved += 1
                    ElseIf PostOL = True Then                           '1      0/5/13 -> -1/4/12   13B7b   1B1b            1 bit
                        SavePostOverlap()
                        BitsSaved += 1
                        Find2ByteMatches = True
                    End If
                Else 'Only if we have more than 45 literals
                    If PreM = True Then
                        SavePreMatch()
                        BitsSaved += 1
                    ElseIf PostM = True Then
                        SavePostMatch()
                        BitsSaved += 1
                    End If
                End If
            Case 1          '2B5b
                '0+1,1+0
                If LitCnt = 1 Then
                    If PreM = True Then                                 '2      1 -> -1                     2B1b            4 bits
                        SavePreMatch()
                        BitsSaved += 4
                    ElseIf PostM = True Then                            '2      1 -> -1                             2B1b    4 bits
                        SavePostMatch()
                        BitsSaved += 4
                    ElseIf (PreOL = True) And (PostOL = True) Then      '1+1    1 -> -1                     1B1b    1B1b    3 bits
                        SavePreOverlap()
                        SavePostOverlap()
                        BitsSaved += 3
                        Find2ByteMatches = True
                    ElseIf PreOL = True Then                            '1      1 -> 0              1B2b    1B1b            2 bits
                        SavePreOverlap()
                        BitsSaved += 2
                    ElseIf PostOL = True Then                           '1      1 -> 0              1B2b            1B1b    2 bits
                        SavePostOverlap()
                        BitsSaved += 2
                        Find2ByteMatches = True
                    End If
                Else    'We have more than 45 literals
                    If PreM = True Then
                        SavePreMatch()
                        BitsSaved += 4
                    ElseIf PostM = True Then
                        SavePreMatch()
                        BitsSaved += 4
                    ElseIf (PreOL = True) And (PostOL = True) Then
                        SavePreOverlap()
                        SavePostOverlap()
                        BitsSaved += 3
                        Find2ByteMatches = True
                    ElseIf (PreM = True) And (PostOL = True) Then
                        SavePreMatch()
                        SavePostOverlap()
                        BitsSaved += 3
                        Find2ByteMatches = True
                    ElseIf (PreOL = True) And (PostM = True) Then
                        SavePreOverlap()
                        SavePostMatch()
                        BitsSaved += 3
                    ElseIf (PreM = True) And (PostM = True) Then
                        SavePreMatch()
                        SavePostMatch()
                        BitsSaved += 3
                    ElseIf PreOL = True Then
                        SavePreOverlap()
                        BitsSaved += 2
                    ElseIf PostOL = True Then
                        SavePostOverlap()
                        BitsSaved += 2
                        Find2ByteMatches = True
                    End If
                End If
            Case 2          '3B5b
                '0+2,1+1
                If LitCnt = 2 Then
                    If (PreM = True) And (PostOL = True) Then           '2+1    2 -> -1                     2B1b    1B1b    3 bits
                        SavePreMatch()
                        SavePostOverlap()
                        BitsSaved += 3
                        Find2ByteMatches = True
                    ElseIf (PreOL = True) And (PostM = True) Then       '1+2    2 -> -1                     1B1b    2B1b    3 bits
                        SavePreOverlap()
                        SavePostMatch()
                        BitsSaved += 3
                    ElseIf PreM = True Then                             '2      2 -> 0                      1B2b    2B1b    2 bits
                        SavePreMatch()
                        BitsSaved += 2
                    ElseIf PostM = True Then                            '2      2 -> 0                      1B2b    2B1b    2 bits
                        SavePostMatch()
                        BitsSaved += 2
                    End If
                Else    'We have more than 45 literals
                    If (PreM = True) And (PostM = True) Then
                        SavePreMatch()
                        SavePostMatch()
                        BitsSaved += 4
                    ElseIf PreM = True Then
                        SavePreMatch()
                        BitsSaved += 3
                    ElseIf PostM = True Then
                        SavePreMatch()
                        BitsSaved += 3
                    ElseIf (PreOL = True) And (PostOL = True) Then
                        SavePreOverlap()
                        SavePostOverlap()
                        BitsSaved += 3
                        Find2ByteMatches = True
                    ElseIf (PreM = True) And (PostOL = True) Then
                        SavePreMatch()
                        SavePostOverlap()
                        BitsSaved += 3
                        Find2ByteMatches = True
                    ElseIf (PreOL = True) And (PostM = True) Then
                        SavePreOverlap()
                        SavePostMatch()
                        BitsSaved += 3
                    End If
                End If
            Case 3          '4B5b, MaxLit+3
                '0+3,1+2
                If (PreM = True) And (PostM = True) Then            '2+2    3 -> -1                     2B1b    2B1b    3 bits
                    SavePreMatch()
                    SavePostMatch()
                    BitsSaved += 3
                ElseIf (PreM = True) And (PostOL = True) Then       '2+1    3 -> 0              1B2B    2B1b    1B1b    1 bit
                    SavePreMatch()
                    SavePostOverlap()
                    BitsSaved += 1
                    Find2ByteMatches = True
                ElseIf (PreOL = True) And (PostM = True) Then       '1+2    3 -> 0              1B2B    1B1b    2B1b    1 bit
                    SavePreOverlap()
                    SavePostMatch()
                    BitsSaved += 1
                End If
            Case 4          '4B5b       8, 16 do not save anything
                If (PreM = True) And (PostM = True) Then            '2+2    4/8/16 -> 0/4/12    1B2b    2B1b    2B1b    1 bit
                    SavePreMatch()
                    SavePostMatch()
                    BitsSaved += 1
                End If
            Case 5, 13      '6B7b, 14B9b
                '0+1,1+0
                If PreOL = True Then                                '1      5/13 -> 4/12        5B5b    1B1b            1 bits
                    SavePreOverlap()    'Check overlap first as it is a short match, faster decompression than mid match
                    BitsSaved += 1
                ElseIf PostOL = True Then                           '1      5/13 -> 4/12        5B5b            1B1b    1 bits
                    SavePostOverlap()
                    BitsSaved += 1
                    Find2ByteMatches = True
                ElseIf PreM = True Then                             '2      5/13 -> 3/11        4b5B    2B1b            1 bits
                    SavePreMatch()
                    BitsSaved += 1
                ElseIf PostM = True Then                            '2      5/13 -> 3/11        4b5B            2B1b    1 bits
                    SavePostMatch()
                    BitsSaved += 1
                End If
            Case 6, 14      '7B7b, 15B9b
                '5/13+1
                If PreM = True Then                                 '2      6/14 -> 4/12        5B5b    2B1b            1 bit
                    SavePreMatch()
                    BitsSaved += 1
                ElseIf PostM = True Then                            '2      6/14 -> 4/12        13B7b   2B1b            1 bit
                    SavePostMatch()
                    BitsSaved += 1
                End If
        End Select

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function PreMatch() As Boolean
        On Error GoTo Err

        'Find 2-byte MidMatches for the first 2 Literal bytes
        PreMatch = False
        'POffset points at the first byte of the new match sequence here
        '     POffset                                                LastPoffset
        '     |                                                       |
        '     v               LitCnt+1                   LastMLen+1   v
        '------|-------------------------------------|-----------------|-----
        '                                          ^^|
        '                                           ||             
        Dim PO = POffset + LitCnt + 1 '-------------+|  Points at the 1st Literal

        If PO > PrgLen Then Exit Function

        CMEnd = PO + 256

        If CMEnd > MatchStart Then CMEnd = MatchStart

        For CmPos = PO + 1 To CMEnd

            If (Prg(PO) = Prg(CmPos)) And (Prg(PO - 1) = Prg(CmPos - 1)) Then

                'PrePOffset = PO
                PreMOffset = CmPos - PO                 'Save offset (1-256)
                PreMatch = True
                Exit Function

            End If

        Next

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function PreOverlap() As Boolean
        On Error GoTo Err

        'Find 2-byte overlapping ShortMatches for the first Literal byte and last Match byte
        PreOverlap = False

        If LastMLen < 2 Then Exit Function

        'POffset points at the first byte of the new match sequence here
        '     POffset                                                LastPoffset
        '     |                                                       |
        '     v               LitCnt+1                   LastMLen+1   v
        '------|-------------------------------------|-----------------|-----
        '                                           ^|^
        '                                            ||             
        Dim PO = POffset + LitCnt + 2 '--------------+|  Points at the last MatchByte

        If PO > PrgLen Then Exit Function

        CMEnd = PO + 64

        If CMEnd > MatchStart Then CMEnd = MatchStart

        For CmPos = PO + 1 To CMEnd

            If (Prg(PO) = Prg(CmPos)) And (Prg(PO - 1) = Prg(CmPos - 1)) Then

                'PreOLPO = PO
                PreOLMO = CmPos - PO                 'Save offset (1-64)
                PreOverlap = True
                Exit Function

            End If

        Next

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function PostMatch() As Boolean
        On Error GoTo Err

        'Find 2-byte MidMatches for the last 2 Literal bytes
        PostMatch = False
        'POffset points at the first byte of the new match sequence here
        '     POffset                                                LastPOffset
        '     |                                                       |
        '     v               LitCnt+1                   LastMLen+1   v
        '------|-------------------------------------|-----------------|-----
        '      |^^                                 ^^|
        '      | |---------+                        ||             
        Dim PO = POffset + 2 'Points at the 2st to last Literal

        If PO > PrgLen Then Exit Function

        CMEnd = PO + 256

        If CMEnd > MatchStart Then CMEnd = MatchStart

        For CmPos = PO + 1 To CMEnd

            If (Prg(PO) = Prg(CmPos)) And (Prg(PO - 1) = Prg(CmPos - 1)) Then

                'PostPOffset = PO
                PostMOffset = CmPos - PO                'Save offset
                PostMatch = True
                Exit Function

            End If

        Next

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function PostOverlap() As Boolean
        On Error GoTo Err

        'Find 2-byte overlapping ShortMatches for the last Literal byte and first Match byte
        PostOverlap = False

        If MLen < 2 Then Exit Function

        'POffset points at the first byte of the new match sequence here
        '     POffset                                                LastPOffset
        '     |                                                       |
        '     v               LitCnt+1                   LastMLen+1   v
        '------|-------------------------------------|-----------------|-----
        '     ^|^                                    |
        '      ||----------+                         |             
        Dim PO = POffset + 1 'Points at the last Literal

        If PO > PrgLen Then Exit Function

        CMEnd = PO + 64

        If CMEnd > MatchStart Then CMEnd = MatchStart

        For CmPos = PO + 1 To CMEnd

            If (Prg(PO) = Prg(CmPos)) And (Prg(PO - 1) = Prg(CmPos - 1)) Then

                'PostOLPO = PO
                PostOLMO = CmPos - PO                'Save offset
                PostOverlap = True
                Exit Function

            End If

        Next

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Sub SavePreMatch()
        On Error GoTo Err

        'We are overwriting the first 2 literal bytes of the last literal sequence with a ShortMatch

        Dim PreByteCnt As Integer = ByteCnt + LitCnt + 1        'Calculate update position in ByteStream

        AddRBits(0, 1)                                          'We need Match Tag

        Buffer(PreByteCnt) = 4                                  'Length is always 1*4
        Buffer(PreByteCnt - 1) = PreMOffset - 1                 'Save Match Offset

        LitCnt -= 2                                             'New Literal Count

        PreM = False

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SavePreOverlap()
        On Error GoTo Err

        'We are overwriting the first literal byte of the last literal sequence with a short MidMatch

        Dim PreByteCnt As Byte

        LastMLen -= 1

        If LastMType = "s" Then                                 'Update previous match as a short match
            Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
            PreByteCnt = LastByteCt - 1
        ElseIf LastMType = "m" Then                             'Update previous match as a mid match
            Buffer(LastByteCt) = LastMLen * 4
            Buffer(LastByteCt - 1) = LastMOffset - 1
            PreByteCnt = LastByteCt - 2
        Else                                                    'Update previous match as a long match
            Buffer(LastByteCt) = LongMatchTag
            Buffer(LastByteCt - 1) = LastMLen
            Buffer(LastByteCt - 2) = LastMOffset - 1
            PreByteCnt = LastByteCt - 3
        End If

        Dim LPO As Integer = LastPOffset

        If (LastMLen <= MaxShortLen) And (LastMType = "m") Then
            For J As Integer = 1 To 64
                Select Case LastMLen
                    Case 1  'MatchLen=2
                        If (Prg(LPO) = Prg(LPO + J)) And (Prg(LPO - 1) = Prg(LPO - 1 + J)) Then
Mid2Short:
                            'Debug.Print("M->S")
                            LastMType = "s"
                            LastMOffset = J - 1
                            Buffer(LastByteCt) = (LastMOffset * 4) + LastMLen
                            For K As Integer = 0 To LitCnt
                                Buffer(LastByteCt - 1 - K) = Buffer(LastByteCt - 2 - K)
                            Next
                            PreByteCnt = LastByteCt - 1
                            ByteCnt += 1
                            BytesSaved += 1
                            Exit For
                        End If
                    Case 2  'MatchLen=3
                        If (Prg(LPO) = Prg(LPO + J)) And (Prg(LPO - 1) = Prg(LPO - 1 + J)) And (Prg(LPO - 2) = Prg(LPO - 2 + J)) Then
                            GoTo Mid2Short
                        End If
                    Case 3  'MatchLen=4
                        If (Prg(LPO) = Prg(LPO + J)) And (Prg(LPO - 1) = Prg(LPO - 1 + J)) And (Prg(LPO - 2) = Prg(LPO - 2 + J)) And (Prg(LPO - 3) = Prg(LPO - 3 + J)) Then
                            GoTo Mid2Short
                        End If
                End Select
            Next
        ElseIf (LastMLen <= MaxMidLen) And (LastMType = "l") Then
            'Debug.Print("L->M")
            LastMType = "m"
            Buffer(LastByteCt) = LastMLen * 4                         'Length of match (#$02-#$3f, cannot be #$00 (end byte), and #$01 - distant selector??)
            Buffer(LastByteCt - 1) = LastMOffset - 1
            For K As Integer = 0 To LitCnt
                Buffer(LastByteCt - 2 - K) = Buffer(LastByteCt - 3 - K)
            Next
            PreByteCnt = LastByteCt - 2
            ByteCnt += 1
            BytesSaved += 1
        End If

        LitCnt -= 1                                             'New Literal Count

        AddRBits(0, 1)                                          'We need Match Tag

        Buffer(PreByteCnt) = ((PreOLMO - 1) * 4) + 1         'Length is always 1

        PreOL = False

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SavePostMatch()
        On Error GoTo Err

        'We are overwriting the last 2 literal bytes of the last literal sequence with a short MidMatch

        Dim PostByteCnt As Integer = ByteCnt + 2                'Calculate update position in ByteStream

        LitCnt -= 2                                             'New Literal Count

        AddLitBits()

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   'If last Literal Length was -1 or Max, we need the Match Tag

        Buffer(PostByteCnt) = 4                                 'Length is always 1*4
        Buffer(PostByteCnt - 1) = PostMOffset - 1               'Save Match Offset

        LitCnt = -1

        LastByteCt = PostByteCnt
        LastMOffset = PostMOffset
        LastMLen = 1
        LastMType = "m"
        LastPOffset = POffset + 2

        PostM = False

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SavePostOverlap()
        On Error GoTo Err

        'We are overwriting the last literal byte of the last literal sequence with a ShortMatch

        Dim PostByteCnt As Integer = ByteCnt + 1         'Calculate update position in ByteStream

        LitCnt -= 1                                     'New Literal Count

        AddLitBits()                                    'Save literal bits

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   'If last Literal Length was -1 or Max, we need the Match Tag

        Buffer(PostByteCnt) = ((PostOLMO - 1) * 4) + 1   'Length is always 1

        LitCnt = -1                                     'reset Literal counter

        LastByteCt = PostByteCnt
        LastMOffset = PostOLMO
        LastMLen = 1
        LastMType = "s"
        LastPOffset = POffset + 1

        PostOL = False

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function LongMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("LongMatch")
        LongMatch = True

        'LONGMATCH

        If LitCnt > -1 Then
            'If there is a PostOverLap then we exit function and start FindMatch from next POffset again`
            If Find2ByteMatches() = True Then Exit Function
        End If

        CalcNeededBits(True)

        If SequenceFits(3, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
            'Add long match here
            AddLitBits()    'First close last literal sequence
            AddLongM()      'Then add Long Match
        Else    'Long match does not fit, check if mid match would fit
            If MaxLen > MaxMidLen Then MaxLen = MaxMidLen           'Adjust MaxLen for MidMatch
            If SequenceFits(2, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                'Add mid match here
                AddLitBits()
                AddMidM()
                GoTo BufferFull
            Else
                If (MaxSLen <> 0) And ((SequenceFits(1, LitCnt, CheckIO(POffset - MaxLen + 1)))) Then
                    'MsgBox("!!")
                    MaxLen = MaxSLen
                    MaxOffset = MaxSOff
                    'Add short match here
                    AddLitBits()
                    AddShortM()
                    GoTo BufferFull
                Else
                    'Match does not fit, check if we can add byte as literal
                    MaxLen = 0
                    CalcNeededBits()
                    If SequenceFits(1, LitCnt + 1, CheckIO(POffset)) Then AddLitByte()        'Add Byte As literal And update LitCnt
                    AddLitBits()
BufferFull:
                    CloseBuff()
                    LongMatch = False    'This will restart FindMatch without advancing POffset
                End If
            End If
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function MidMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("MidMatch")
        MidMatch = True

        'MIDMATCH

        If LitCnt > -1 Then
            'If there is a PostOverLap then we exit function and start FindMatch from next POffset again`
            If Find2ByteMatches() = True Then Exit Function
        End If

        CalcNeededBits(True)

        'Check if Mid Match fits
        If SequenceFits(2, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
            AddLitBits()        'First close last literal sequence, will not do anything if LitCnt=-1

            'If no literals and MidMatch is preceded by another MidMatch and total length = 6-12 bytes
            'then see if we can replace them with 3 ShortMatches (3B3b vs 4B2b) to save 7 bits
            If (LitCnt = -1) And (LastMType = "m") And (MaxLen + 1 + LastMLen + 1 <= 12) And (MaxLen + 1 + LastMLen + 1 >= 6) Then
                '--------------------------------------------------------------
                'If ReplaceMidWithShort() = True Then Exit Function
                '--------------------------------------------------------------
            End If

            AddMidM()           'Then add Mid Match
        Else
            If (MaxSLen <> 0) And ((SequenceFits(1, LitCnt, CheckIO(POffset - MaxLen + 1)))) Then
                'MsgBox("!")
                MaxLen = MaxSLen
                MaxOffset = MaxSOff
                AddLitBits()
                AddShortM()
                GoTo BufferFull
            Else
                'Match does not fit, check if we can add byte as literal
                MaxLen = 0
                CalcNeededBits()
                If SequenceFits(1, LitCnt + 1, CheckIO(POffset)) Then AddLitByte()        'Add Byte As literal And update LitCnt
                AddLitBits()
BufferFull:
                CloseBuff()
                MidMatch = False    'This will restart FindMatch without advancing POffset
            End If
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function ShortMatch() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("ShortMatch")
        ShortMatch = True

        'SHORTMATCH

        If LitCnt > -1 Then
            'If there is a PostOverLap then we exit function and start FindMatch from next POffset again
            If Find2ByteMatches() = True Then Exit Function
        End If

        CalcNeededBits(True)

        'We have already reserved space for literal bits, check is match fits
        If SequenceFits(1, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
            'Yes, we have space
            AddLitBits()    'First close last literal sequence
            AddShortM()       'Then add Near Byte
        Else
            'Match does not fit, check if we can add byte as literal
            MaxLen = 0
            CalcNeededBits()
            If SequenceFits(1, LitCnt + 1, CheckIO(POffset)) Then AddLitByte()           'Add byte as literal and update LitCnt

            AddLitBits()
            CloseBuff()
            ShortMatch = False    'This will restart FindMatch without advancing POffset
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function Literal() As Boolean
        On Error GoTo Err

        If DoDebug Then Debug.Print("Literal")
        Literal = True

        'LITERAL

        CalcNeededBits()

        If SequenceFits(1, LitCnt + 1, CheckIO(POffset)) = True Then 'We are adding 1 literal byte and 1+5 bits
            'We have space in the buffer for this byte
            AddLitByte()
        Else    'No space for this literal byte, close sequence
            'If LitCnt > -1 Then AddLitBits()
            AddLitBits()
            CloseBuff()
            Literal = False    'Goto Restart
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        Literal = False

    End Function

    Private Sub CalcNeededBits(Optional ForMatch As Boolean = False)
        On Error GoTo Err

        If ForMatch = True Then
            'The next sequence is a match

            Bits = Fix(LitCnt / (MaxLit + 1)) * 9

            Select Case LitCnt Mod (MaxLit + 1)
                Case -1
                    Bits = 0 + 1            'No Literals, we need 1 match tag bit
                Case 0
                    Bits += 2 + 0           '2 lit tag bits
                Case 1 To 4
                    Bits += 3 + 2 + 0       '3 lit tag bits + 2 lit sequence length bits
                Case 5 To 12
                    Bits += 4 + 3 + 0       '4 lit tag bits + 3 lit sequence length bits
                Case 13 To MaxLit - 1
                    Bits += 4 + 5 + 0       '4 lit tag bits + 5 lit sequence length bits
                Case MaxLit
                    Bits += 4 + 5 + 1       '4 lit tag bits + 5 lit sequence length bits + 1 match tag bit
            End Select
        Else
            'The next byte is a literal byte

            Bits = Fix((LitCnt + 1) / (MaxLit + 1)) * 9

            Select Case (LitCnt + 1) Mod (MaxLit + 1)
                Case 0
                    Bits += 2 + 0           '2 lit tag bits
                Case 1 To 4
                    Bits += 3 + 2 + 0       '3 lit tag bits + 2 lit sequence length bits
                Case 5 To 12
                    Bits += 4 + 3 + 0       '4 lit tag bits + 3 lit sequence length bits
                Case 13 To MaxLit - 1
                    Bits += 4 + 5 + 0       '4 lit tag bits + 5 lit sequence length bits
                Case MaxLit
                    Bits += 4 + 5 + 1       '4 lit tag bits + 5 lit sequence bits + 1 match tag bit for the following match
            End Select
        End If

        'If LitCnt >= MaxLit Then MsgBox(Bits.ToString + vbNewLine + LitCnt.ToString)

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function SequenceFits(ByteLen As Integer, Literals As Integer, Optional SequenceUnderIO As Integer = 0) As Boolean
        On Error GoTo Err

        'Calculate total bit count in buffer from ByteCnt, BitCnt, and BitPos
        Dim BitsInBuffer As Integer = ((255 - ByteCnt + BitCnt) * 8) + 16 - BitPos

        'Add 1 Close Byte, and 1 more if this is the first sequence in the block that goes under IO
        Dim NeededBytes As Integer = ByteLen + 1 + If((BlockUnderIO = 0) And (SequenceUnderIO = 1), 1, 0)

        'Close Bit, not needed most of the time, cannot overlap with Close Byte!!!
        'If we are trying to add a match sequence here then we need an additional match tag close bit for the close match byte
        'which is not needed if this is a literal sequence as CalcLitBits automatically adds a match bit for the next match if needed
        Dim CloseBit As Integer = If(MaxLen > 0, 1, 0)

        Dim BitLen As Integer = Bits + (NeededBytes * 8) + CloseBit

        If BitsInBuffer + BitLen <= 2048 Then
            SequenceFits = True
            'Data will fit
            If (BlockUnderIO = 0) And (SequenceUnderIO = 1) Then
                'This is the first sequence in the block that will go UIO, so lets update the buffer to include the IO flag
                For I As Integer = ByteCnt To AdHiPos            'Move all data to the left in buffer, including AdHi
                    Buffer(I - 1) = Buffer(I)
                Next
                Buffer(AdHiPos) = 0                             'IO Flag to previous AdHi Position
                ByteCnt -= 1                                    'Update ByteCt to next empty position in buffer
                LastByteCt -= 1                                 'Last Match pointer also needs to be updated (BUG FIX - REPORTED BY RAISTLIN/G*P)
                AdHiPos -= 1                                    'Update AdHi Position in Buffer
                BlockUnderIO = 1                                'Set BlockUnderIO Flag
            End If
        Else
            SequenceFits = False
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Sub AddLongM()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddLongM")

        TotMatch += 1

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCnt) = LongMatchTag                   'Long Match Flag = &HF8
        Buffer(ByteCnt - 1) = MaxLen
        Buffer(ByteCnt - 2) = MaxOffset - 1
        ByteCnt -= 3

        POffset -= MaxLen

        LitCnt = -1

        LM += 1

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddMidM()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddMidM")

        TotMatch += 1

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCnt) = MaxLen * 4                         'Length of match (#$02-#$3f, cannot be #$00 (end byte), and #$01 - distant selector??)
        Buffer(ByteCnt - 1) = MaxOffset - 1
        ByteCnt -= 2

        POffset -= MaxLen

        LitCnt = -1

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddShortM()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddShortM")

        TotMatch += 1

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCnt) = ((MaxOffset - 1) * 4) + MaxLen
        ByteCnt -= 1

        POffset -= MaxLen

        LitCnt = -1

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddLitByte()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddLitByte")

        LitCnt += 1             'Increase Literal Counter

        Buffer(ByteCnt) = Prg(POffset)   'Add byte to Byte Stream
        ByteCnt -= 1                     'Update Byte Position Counter

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddLitBits()
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddLitBits")

        If LitCnt = -1 Then Exit Sub

        TotLit += Int(LitCnt / (MaxLit + 1)) + 1

        For I As Integer = 1 To Fix(LitCnt / (MaxLit + 1))
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

        Dim Lits As Integer = LitCnt Mod (MaxLit + 1)

        Select Case Lits
            Case 0
                AddRBits(0, 1)              'Add Literal Length Selector 0	- read no more bits
            Case 1 To 4
                AddRBits(2, 2)              'Add Literal Length Selector 10 - read 2 more bits
                AddRBits(Lits - 1, 2)     'Add Literal Length: 00-03, 2 bits	-> 1000 00xx when read
            Case 5 To 12
                AddRBits(6, 3)              'Add Literal Length Selector 110 - read 3 more bits
                AddRBits(Lits - 5, 3)     'Add Literal Length: 00-07, 3 bits	-> 1000 1xxx when read
            Case 13 To MaxLit
                AddRBits(7, 3)              'Add Literal Length Selector 111 - read 5 more bits
                AddRBits(Lits - 13, 5)    'Add Literal Length: 00-1f, 5 bits	-> 101x xxxx when read
        End Select

        'DO NOT RESET LitCnt HERE!!!

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddRBits(Bit As Integer, BCnt As Byte)
        On Error GoTo Err

        If DoDebug Then Debug.Print("AddRBits:" + vbTab + Hex(Bit).ToString + vbTab + BCnt.ToString)

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
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub CloseBuff()  'CHANGE TO PUBLIC
        On Error GoTo Err

        If DoDebug Then Debug.Print("CloseBuff")

        Buffer(ByteCnt) = EndTag                'Technically not needed, defaulte value is #$00 anyway
        Buffer(0) = Buffer(0) And &H7F          'Delete Compression Bit (Default (i.e. compressed) value is 0)

        'FIND UNCOMPRESSIBLE BLOCKS (only applies to first file in buffer in which case LastByte points at the very first data byte in stream)

        'If (MatchStart - POffset <= LastByte) And (MatchStart > LastByte - 1) And (NextFileInBuffer = False) Then
        If (MatchStart - POffset <= LastByte) And (MatchStart > LastByte - 1) And (NextFileInBuffer = False) Then
            'MsgBox((LastByte - (MatchStart - POffset)).ToString)
            'Less than 252/253 bytes        AND  not the end of File     AND       No other files in this buffer
            LastByte = AdLoPos - 2

            'Check uncompressed Block IO Status
            'If (CheckIO(MatchStart) Or CheckIO(MatchStart - (LastByte - 1)) = 1) And (FileUnderIO = True) Then
            If CheckIO(MatchStart) Or CheckIO(MatchStart - (LastByte - 1)) = 1 Then
                'If the block will be UIO than only (Lastbyte-1) bytes will fit,
                'So we only need to check that many bytes
                Buffer(AdLoPos - 1) = 0 'Set IO Flag
                AdHiPos = AdLoPos - 2   'Updae AdHiPos
                LastByte = AdHiPos - 1  'Update LastByte
                'POffset += 1
                'ElseIf (CheckIO(MatchStart - LastByte) = 1) And (FileUnderIO = True) Then
            ElseIf CheckIO(MatchStart - LastByte) = 1 Then
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
            ByteCnt = 1

        End If

        BlockCnt += 1
        BufferCnt += 1
        UpdateByteStream()

        ResetBuffer()               'Resets buffer variables

        NextFileInBuffer = False            'Reset Next File flag

        If POffset < 0 Then Exit Sub 'We have reached the end of the file -> exit

        'If we have not reached the end of the file, then update buffer

        BlockUnderIO = CheckIO(POffset)          'Check if last byte of prg could go under IO

        Buffer(ByteCnt) = (PrgAdd + POffset) Mod 256
        AdLoPos = ByteCnt

        If BlockUnderIO = 1 Then ByteCnt -= 1

        Buffer(ByteCnt - 1) = Int((PrgAdd + POffset) / 256) Mod 256
        AdHiPos = ByteCnt - 1
        ByteCnt -= 2
        LastByte = ByteCnt              'LastByte = the first byte of the ByteStream after and Address Bytes (253 or 252 with blockCnt)

        Buffer(ByteCnt) = Prg(POffset)  'Copy First Lit Byte to Buffer
        ByteCnt -= 1                    'Update Byte Pos Counter
        LastMS = MatchStart
        MatchStart = POffset            'Update Match Start Position
        POffset -= 1                    'Next Byte in Prg
        If POffset < 0 Then POffset = 0 'Unless it is <0		NEEDED DO NOT DELETE!!!
        LastPO = POffset
        LitCnt = 0                      'Literal counter has 1 value

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub UpdateByteStream()   'THIS IS ALSO USED BY LZ4+RLE!!!
        On Error GoTo Err

        ReDim Preserve ByteSt(BufferCnt * 256 - 1)

        For I = 0 To 255
            ByteSt((BufferCnt - 1) * 256 + I) = Buffer(I)
        Next

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SaveLastMatch()
        On Error GoTo Err

        LastByteCt = ByteCnt
        LastMOffset = MaxOffset
        LastMLen = MaxLen
        LastMType = MaxType
        LastPOffset = POffset

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub ResetBuffer() 'CHANGE TO PUBLIC
        On Error GoTo Err

        ReDim Buffer(255)       'New empty buffer

        'Initialize variables

        FilesInBuffer = 1

        If BufferCnt = 0 Then
            ByteCnt = 254        'First file in part, last byte of first block will be BlockCnt

        Else
            ByteCnt = 255        'All other blocks:	Bytestream start at the end of the buffer
        End If

        BitCnt = 0               'Bitstream always starts at the beginning of the buffer

        BitPos = 15             'Reset Bit Position Counter (counts 16 bits backwards: 15-0)

        'DO NOT RESET LitCnt HERE!!! It is needed for match tag check

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Function CheckIO(Offset As Integer, Optional NextFileUnderIO As Integer = -1) As Integer
        On Error GoTo Err

        Offset += PrgAdd

        If Offset < 256 Then       'Are we loading to the Zero Page? If yes, we need to signal it by adding IO Flag
            CheckIO = 1
        ElseIf NextFileUnderIO > -1 Then
            CheckIO = If((Offset >= &HD000) And (Offset <= &HDFFF) And (NextFileUnderIO = 1), 1, 0)
        Else
            CheckIO = If((Offset >= &HD000) And (Offset <= &HDFFF) And (FileUnderIO = True), 1, 0)
        End If

        Exit Function
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Public Function FinishPart(Optional NextFileIO As Integer = 0, Optional LastPartOnDisk As Boolean = False) As Boolean
        On Error GoTo Err

        'MatchStart = LastMS
        'CompressPart(LastPO)

        FinishPart = True

        If NewBlock = True Then GoTo NewB   'The part will start in a new block

        'ADDS NEW PART TAG (Long Match Tag + End Tag) TO THE END OF THE PART, AND RESERVES LAST BYTE IN BUFFER FOR BLOCK COUNT
        Dim Bytes, Bits As Integer

        '"Sprite bug"
        'Compression bug related to the transitional block, reported by Raistlin/G*P - FIXED
        'Fix: include next file's I/O status in calculation of needed bytes
        Bytes = 6 + NextFileIO 'BYTES NEEDED: BlockCnt + Long Match Tag + End Tag + AdLo + AdHi + 1st Literal (ByC=7 if BlockUnderIO=true - checked at SequenceFits)
        'If LastPartOnDisk = True Then Bytes += 1
        Bits = If((LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit), 1, 0)   'Calculate whether Match Bit is needed for new part

        If SequenceFits(Bytes, LitCnt) Then

            'Buffer has enough space for New Part Tag and New Part Info and first Literal byte (and IO flag if needed)

            If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)  'Add Match Selector Bit only if needed
NextPart:

            FilesInBuffer += 1  'There is going to be more than 1 file in the buffer

            If (BufferCnt > 0) And (FilesInBuffer = 2) Then         'Reserve last byte in buffer for Block Count...
                For I = ByteCnt + 1 To 255                          '... only once, when the 2nd file is added to the same buffer
                    Buffer(I - 1) = Buffer(I)
                Next
                ByteCnt -= 1
                Buffer(255) = 1                                     'Last byte reserved for BlockCnt
            End If

            Buffer(ByteCnt) = LongMatchTag                          'Then add New File Match Tag
            Buffer(ByteCnt - 1) = EndTag
            ByteCnt -= 2

            If LastPartOnDisk = True Then
                Buffer(ByteCnt) = ByteCnt - 2   'Finish disk with a dummy literal byte that overwrites itself to reset LastX for next disk side
                Buffer(ByteCnt - 1) = &H3       'New address is the next byte in buffer
                Buffer(ByteCnt - 2) = &H0       'Dummy $00 Literal that overwrites itself
                LitCnt = 0                      'One (dummy) literal
                'AddLitBits()                   'NOT NEEDED, WE ARE IN THE MIDDLE OF THE BUFFER, 1ST BIT NEEDS TO BE OMITTED
                AddRBits(0, 1)                  'ADD 2ND BIT SEPARATELY (0-BIT, TECHNCALLY, THIS IS NOT NEEDED)
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
NewB:            'Next File Info does not fit, so close buffer
            CloseBuff()
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
                MsgBox("Part " + If(CompressPartFromEditor = True, PartCnt + 1, PartCnt).ToString + " would need " + LastBlockCnt.ToString + " blocks on the disk." + vbNewLine + vbNewLine + "Parts cannot be larger than 255 blocks!", vbOKOnly + vbCritical, "Part exceeds 255-block limit!")
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
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")
NoGo:
        FinishPart = False

    End Function

    Public Sub FinishFile()
        On Error GoTo Err

        'ADDS NEXT FILE TAG TO BUFFER

        Dim Bytes, Bits As Integer

        '4 bytes and 0-1 bits needed for NextFileTag, Address Bytes and first Lit byte (+1 more if UIO)
        Bytes = 4 'BYTES NEEDED: End Tag + AdLo + AdHi + 1st Literal (ByC=5 only if BlockUnderIO=true - checked at SequenceFits()
        Bits = If((LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit), 1, 0)   'Calculate whether Match Bit is needed for new file

        If SequenceFits(Bytes, LitCnt, CheckIO(PrgLen - 1)) Then

            'Buffer has enough space for New File Match Tag and New File Info and first Literal byte (and IO flag if needed)

            If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)  'Add Match Selector Bit only if needed

            Buffer(ByteCnt) = NextFileTag                           'Then add New File Match Tag
            ByteCnt -= 1
        Else
            'Next File Info does not fit, so close buffer
            CloseBuff()
        End If

        Exit Sub
Err:
        ErrCode = Err.Number
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

End Module
