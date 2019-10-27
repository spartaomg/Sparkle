Friend Module ModNewLZ
    Private CMStart, CMEnd, POffset, MLen As Integer
    Private ReadOnly MaxLongLen As Integer = 254    'Cannot be 255, there is an INY in the decompression ASM code, and that would make YR=#$00
    Private ReadOnly MaxMidLen As Integer = 61      'Cannot be more than 61 because 62=LongMatchTag, 63=NextFileTage
    Private ReadOnly MaxShortLen As Integer = 3     '1-3, cannot be 0 because it is preserved for EndTag
    Private ReadOnly DoDebug As Boolean = False
    Private ReadOnly MaxLit As Integer = 1 + 4 + 8 + 32 - 1  '=44

    Private ReadOnly LongMatchTag As Byte = &HF8    'Could be changed to &H00, but this is more economical
    Private ReadOnly NextFileTag As Byte = &HFC
    Private ReadOnly EndTag As Byte = 0             'Could be changed to &HF8, but this is more economical (Number of EndTags > Number of LongMatchTags)

    Private FirstOfNext As Boolean = False          'If true, this is the first block of next file in same buffer, Lit Selector Bit NOT NEEEDED
    Private NextFile As Boolean = False             'Indicates whether the next file is added to the same buffer

    Private BlockUnderIO As Integer = 0
    Private AdLoPos As Byte, AdHiPos As Byte

    Private PreM As Boolean = False
    Private PostM As Boolean = False
    Private PreOL As Boolean = False
    Private PostOL As Boolean = False
    Private AddPostM As Boolean = False
    Private AddPostOL As Boolean = False

    Public Sub NewLZ(PN As Object, Optional FA As String = "", Optional FO As String = "", Optional FL As String = "", Optional FUIO As Boolean = False)
        'On Error GoTo Err

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

        LastByte = ByteCnt   'ByteCt points at the next empty byte in buffer

NextBuffer:
        BlockUnderIO = 0       'Reset BlockUnderIO Flag

        If ((BlockCnt = 0) And (ByteCnt = 254)) Or (ByteCnt = 255) Then

            'First Block, ByteCt starts at 254, all other blocks: 255

            'Buffer is empty: we do not need ByC and BiC

            FirstOfNext = False                                         'First block in buffer, Lit Selector Bit is needed(?)
            NextFile = False                                            'This is the first file that is being added to an empty buffer
FileAddress:

            Buffer(ByteCnt) = (PrgAdd + PrgLen - 1) Mod 256              'Add Address Info
            AdLoPos = ByteCnt

            'If (FileUnderIO = True) And (CheckIO(PrgLen - 1) = 1) Then  'Check if last byte of block is under IO
            If CheckIO(PrgLen - 1) = 1 Then                             'Check if last byte of block is under IO
                BlockUnderIO = 1                                        'Yes, set BUIO flag
                ByteCnt -= 1                                            'And skip 1 byte (=0) for IO Flag
            End If

            Buffer(ByteCnt - 1) = Int((PrgAdd + PrgLen - 1) / 256)
            AdHiPos = ByteCnt - 1

            ByteCnt -= 2
            LitCnt = -1                                                 'Reset LitCnt here

        Else
            'Buffer is not empty, we need New File Match Tag, check if we have enough space left in open buffer

            '4 bytes and 0-1 bits needed for NextFileTag, Address Bytes and first Lit byte (+1 more if UIO)

            ByC = 4                                                     'ByC=5 only if BlockUnderIO=true - checked at DataFits()
            IIf((LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit), BiC = 1, BiC = 0)   'Calculate number of bits needed for Match Tag for new file

            If DataFits(ByC, BiC, LitCnt, CheckIO(PrgLen - 1)) Then

                'Buffer has enough space for New File Match Tag and New File Info and first Literal byte (and IO flag if needed)

                If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)  'Add Match Selector Bit only if needed
                Buffer(ByteCnt) = NextFileTag                               'Then add New File Match Tag
                ByteCnt -= 1
                FirstOfNext = True                                      'First block of next file in same buffer, Lit Selector Bit NOT NEEEDED
                NextFile = True                                         'Next file is being added to buffer that already has data
                GoTo FileAddress                                        'Then add Address Info
            Else
                'Next File Info does not fit, so close buffer
                CloseBuff()                                             'Then close and reset Buffer
                GoTo NextBuffer                                         'Then add Address Info to new buffer
            End If
        End If

        LastByte = ByteCnt                       'The first byte of the ByteStream after (BlockCnt and IO Flag and) Address Bytes (251..253)

        'BufferCt = 0                           'First buffer of new file
        MatchStart = PrgLen - 1    '2
        Buffer(ByteCnt) = Prg(PrgLen - 1)        'Add last byte of Prg to buffer
        LitCnt += 1
        ByteCnt -= 1

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


    Private Function FindMatch(Optional PO As Integer = -1) As Boolean
        'FINDS MATCHES AND IDENTIFIES THE ONE THAT SAVES THE MOST

        'On Error GoTo Err

        If DoDebug Then Debug.Print("FindMatch")

        If PO = -1 Then PO = POffset

        'CMStart = If(POffset + 256 > MatchStart, MatchStart, POffset + 256)     '255? vs 256?

        CMStart = PO + 256

        If CMStart > MatchStart Then CMStart = MatchStart


        '       PO                            CMPos (CMStart=PO+256 vs MatchStart)
        '    +++P|C...    <-      <-     <-+++C|
        '--------|-----------------------------|-------
        '       0|123456789A...            ...FF

        MatchCnt = 0
        ReDim MatchOffset(0), MatchLen(0), MatchSave(0), MatchType(0)
        MLen = -1

        For CmPos = CMStart To PO + 1 Step -1
NextByte:
            MLen += 1

            If PO - MLen > -1 Then
                If Prg(PO - MLen) = Prg(CmPos - MLen) And (MLen <= MaxLongLen) = True Then GoTo NextByte
            End If

            'Calculate length of matches
            If MLen > 1 Then
                MLen -= 1
                If (CmPos - PO > 64) And (MLen = 1) Then
                    'Exclude 2-byte mid matches
                Else
                    'Save it to Match List
                    MatchCnt += 1
                    ReDim Preserve MatchOffset(MatchCnt)
                    ReDim Preserve MatchLen(MatchCnt)
                    ReDim Preserve MatchSave(MatchCnt)
                    ReDim Preserve MatchType(MatchCnt)
                    MatchOffset(MatchCnt) = CmPos - PO 'Save offset
                    MatchLen(MatchCnt) = MLen               'Save len

                    If MLen > MaxMidLen Then                       'Calculate number of saved bytes and store it in MatchSave
                        'LongMatch (MLen=62-254)
                        MatchSave(MatchCnt) = MLen - 2
                        MatchType(MatchCnt) = "l"
                    ElseIf MLen > MaxShortLen Then
                        'MidMatch (MLen=04-61)
                        MatchSave(MatchCnt) = MLen - 1
                        MatchType(MatchCnt) = "m"
                    Else    'MLen=01-03
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
            Next

            'Select Case MaxType
            'Case "s"
            'MatchBytes = 1
            'Case "m"
            'MatchBytes = 2
            'Case "l"
            'MatchBytes = 3
            'End Select
            'Select Case LitCnt Mod (MaxLit + 1)
            'Case -1, MaxLit
            'MatchBits = 1            'No Literals, we need 1 match tag bit
            'Case Else
            'MatchBits = 0
            'End Select

            If MaxSave = 0 Then
                FindMatch = False
            Else
                FindMatch = True
            End If
        Else
            FindMatch = False
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    'Private Sub FindPrePostMatches()

    'CalcOrigNeeds()     'Calculates last match, literals and new match space occupation

    'FindPreMatch()      'Checks if a match can be found at the beginning of literals that wound decrease the number of literals

    'FindPostMatch()     'Checks if a match can be found at the end of literals that would decrease the number of literals

    'CalcAndUpdate()     'Calculates potential savings and updates byte and bit streams

    'End Sub

    Private Function Find2ByteMatches() As Boolean

        Find2ByteMatches = False

        PreOL = False
        PostOL = False
        PreM = False
        PostM = False

        If LitCnt > 0 Then
            PreM = PreMatch()        'Must be 2-byte MidMatch
            PostM = PostMatch()      'Must be 2-byte MidMatch
        End If
        PreOL = PreOverlap()         'Must be 2-byte ShortMatch
        PostOL = PostOverlap()       'Must be 2-byte ShortMatch

        AddPostM = False
        AddPostOL = False

        'RangeMins:  0,1,5,13
        Select Case LitCnt Mod (MaxLit + 1)
            Case 0 ', 5, 13   '1B2b, 6B7b, 14B9b
                '0/5/13+0                                                                       Lits    Pre     Post    Saved
                If PreOL = True Then                                '1      0/5/13 -> -1/4/12   5B5b    1B1b            1 bit
                    SavePreOverlap()
                    BitsSaved += 1
                ElseIf PostOL = True Then                           '1      0/5/13 -> -1/4/12   13B7b   1B1b            1 bit
                    SavePostOverlap()
                    BitsSaved += 1
                    Find2ByteMatches = True
                End If
            Case 1          '2B5b
                '0+1,1+0
                If PreM = True Then                                 '2      1 -> -1                     2B1b            4 bits
                    SavePreMatch()
                    BitsSaved += 4
                ElseIf PostM = True Then                            '2      1 -> -1                     2B1b            4 bits
                    SavePostMatch()
                    BitsSaved += 4
                ElseIf (PreOL = True) And (PostOL = True) Then      '1+1    1 -> -1                     2B2b            3 bits
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
            Case 2          '3B5b
                '0+2,1+1
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
            Case 3          '4B5b
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
            Case 4 ', 8, 16   '5B5b                                 9B7b, 16B9b - THE LATTER TWO WILL NOT SAVE ANYTHING
                '1/5/13+3
                If (PreM = True) And (PostM = True) Then            '2+2    4/8/16 -> 0/4/12    1B2b    2B1b    2B1b    1 bit
                    '                                                                           5B5b    2B1b    2B1b    0 bit
                    '                                                                           12B7b   2B1b    2B1b    0 bit
                    SavePreMatch()
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
                    '''''ElseIf (PreOL = True) And (PostOL = True) Then '1+1    6/14 -> 4/12        5B5b    1B1b    1B1b    0 bit
                End If
                ''Case 7, 15      '8B7b, 16B9b    'THIS WILL NOT SAVE ANYTHING
                ''5/13+2
                ''If (PreM = True) And (PostOL = True) Then           '2+1    7/15 -> 4/12        5B5b    2B1b    1B1b    0 bit
                ''ElseIf (PreOL = True) And (PostM = True) Then       '1+2    7/15 -> 4/12        13B7b   1B1b    2B1b    0 bit
                ''End If
        End Select

    End Function

    Private Function PreMatch() As Boolean
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

    End Function

    Private Function PreOverlap() As Boolean
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

    End Function

    Private Function PostMatch() As Boolean
        'Find 2-byte MidMatches for the lastst 2 Literal bytes
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

    End Function

    Private Function PostOverlap() As Boolean
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

    End Function

    Private Sub SavePreMatch()
        'We are overwriting the first 2 literal bytes of the last literal sequence with a short MidMatch

        Dim PreByteCnt As Integer = ByteCnt + LitCnt + 1        'Calculate update position in ByteStream

        AddRBits(0, 1)                                          'We need Match Tag

        Buffer(PreByteCnt) = 4                                  'Length is always 1*4
        Buffer(PreByteCnt - 1) = PreMOffset - 1                 'Save Match Offset

        LitCnt -= 2                                             'New Literal Count

        PreM = False

    End Sub

    Private Sub SavePreOverlap()
        'We are overwriting the first literal byte of the last literal sequence with a short MidMatch

        Dim PreByteCnt As Byte

        LastMLen -= 1

        If LastMType = "s" Then   'Update previous match as a short match
            Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
            PreByteCnt = LastByteCt - 1
        ElseIf LastMType = "m" Then                           'Update previous match as a mid match
            Buffer(LastByteCt) = LastMLen * 4
            Buffer(LastByteCt - 1) = LastMOffset - 1
            PreByteCnt = LastByteCt - 2
        Else                                                        'Update previous match as a long match
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
            Debug.Print("L->M")
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

    End Sub

    Private Sub SavePostMatch()
        'We are overwriting the last 2 literal bytes of the last literal sequence with a short MidMatch

        Dim PostByteCnt As Integer = ByteCnt + 2                'Calculate update position in ByteStream

        LitCnt -= 2                                             'New Literal Count

        AddLitBits()

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   'If last Literal Length was -1 or Max, we need the Match Tag

        Buffer(PostByteCnt) = 4                                 'Length is always 1*4
        Buffer(PostByteCnt - 1) = PostMOffset - 1               'Save Match Offset

        LitCnt = -1

        PostM = False

    End Sub

    Private Sub SavePostOverlap()
        'We are overwriting the last literal byte of the last literal sequence with a short MidMatch

        Dim PreByteCnt As Integer = ByteCnt + 1         'Calculate update position in ByteStream

        LitCnt -= 1                                     'New Literal Count

        AddLitBits()                                    'Save literal bits

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   'If last Literal Length was -1 or Max, we need the Match Tag

        Buffer(PreByteCnt) = ((PostOLMO - 1) * 4) + 1   'Length is always 1

        'MaxLen -= 1                                     'Next Match length will be one less 2/2 the overlap
        'POffset -= 1                                   'We have also used the byte under POffset, so POffset is also moved
        LitCnt = -1                                     'reset Literal counter

        PostOL = False

    End Sub

    'Private Sub CalcAndUpdate()

    'NewBytes = LastBytes + PreBytes + PostBytes + MatchBytes + LitBytes + 1
    'NewBits = LastBits + PreBits + PostBits + MatchBits

    'CalcLitBits()

    'NewBits += LitBits

    'Debug.Print("Bytes:" + vbTab + OrigBytes.ToString + vbTab + "Bits:" + vbTab + OrigBits.ToString + vbTab + "NBytes:" + vbTab + NewBytes.ToString + vbTab + "NBits:" + vbTab + NewBits.ToString)

    'End Sub

    'Private Sub CalcOrigNeeds()

    'LitBytes = LitCnt

    'OrigBytes = LastBytes + MatchBytes + LitCnt + 1
    'OrigBits = LastBits + MatchBits

    'CalcLitbits()

    'OrigBits += LitBits

    'End Sub

    'Private Function FindPostMatch(Optional PO As Integer = -1) As Boolean
    'FindPostMatch = False

    'Dim MO, ML As Integer

    'For I As Integer = POffset + 2 To MatchStart - 1
    'If (Prg(POffset + 1) = Prg(I)) And (Prg(POffset + 2) = Prg(I + 1)) And (I < POffset + 1 + 256) Then
    'MO = MaxOffset
    'ML = MaxLen
    'PostMOffset = I - (POffset + 1)
    'PostMLen = 1
    'PostMType = "m"
    'PostPOffset += MaxLen + 1

    ''LitBytes -= 2
    ''ByteCnt += 2
    ''AddLitBits()
    ''AddMidM()
    ''SM2 += 1
    ''MaxOffset = MO
    ''MaxLen = ML
    ''POffset -= 1
    'FindPostMatch = True
    'Exit Function
    'End If
    'Next

    'End Function

    'Private Function FindPreMatch(Optional PO As Integer = -1) As Boolean
    'FindPreMatch = False

    'Dim NeededMLen As Integer = (LitCnt Mod MaxLit + 1) + 2     'LitCnt=0 -> 2 vs LitCnt=1 -> 3

    'If PO = -1 Then PO = POffset + LitCnt + 1 + 1

    'CMEnd = PO + 256

    'If CMEnd > MatchStart Then CMEnd = MatchStart


    '      PO|CMPOs                        CMEnd (=PO+256 vs MatchStart)
    '     ++P|C    ->      ->      ->     C|
    '--------|-----------------------------|-------
    '       0|123456789A...            ...FF

    'PreBytes = 0 : PreBits = 0
    'PreMOffset = 0 : PreMLen = 0
    'PrePOffset = PO
    'PreMType = ""

    'MLen = -1

    'For CmPos = PO + 1 To CMEnd
    'NextByte:
    'MLen += 1

    'If PO - MLen > -1 Then
    'If Prg(PO - MLen) = Prg(CmPos - MLen) And (MLen < NeededMLen) = True Then GoTo NextByte
    'End If

    'If MLen = NeededMLen Then
    ''Calculate length of matches
    'If (CmPos - PO > 64) And (MLen = 2) And (NeededMLen = 2) Then
    ''Exclude 2-byte mid matches if we need 2-byte matches
    'MLen = -1   'Reset MLen and continue search for match
    'Else
    'MLen -= 1   'Match found, finish search
    'PreMOffset = CmPos - PO                 'Save offset
    'PreMLen = MLen                          'Save len

    'If PreMOffset > 64 Then                       'Calculate number of saved bytes and store it in MatchSave
    ''MatchSave(MatchCnt) = MLen - 2
    'PreBytes = 2
    'Else
    'PreBytes = 1
    'End If
    'PreBits = 1
    'FindPreMatch = True
    'Exit For
    'End If
    'End If
    'Next

    'If FindPreMatch = False Then Exit Function

    'LitBytes = LitBytes - (NeededMLen - 1)
    'CalcLitbits()

    'RecalcLastMatch(LastPOffset)

    'End Function

    'Private Sub RecalcLastMatch(PO As Integer)
    ''Check if short midmatches could be converted to shortmatches (find new match)
    'LastMLen -= 1
    'If (LastMOffset > 64) And (LastMLen <= MaxShortLen) Then
    'For J As Integer = 1 To 64
    'Select Case LastMLen
    'Case 1  'MatchLen=2
    'If (Prg(PO + 2) = Prg(PO + 2 + J)) And (Prg(PO + 3) = Prg(PO + 3 + J)) Then
    'Mid2Short:
    'LastMOffset = J - 1
    'LastMType = "s"
    'LastBytes = 1
    'Exit For
    'End If
    'Case 2  'MatchLen=3
    'If (Prg(PO + 2) = Prg(PO + 2 + J)) And (Prg(PO + 3) = Prg(PO + 3 + J)) And (Prg(PO + 4) = Prg(PO + 4 + J)) Then
    'GoTo Mid2Short
    'End If
    'Case 3  'MatchLen=4
    'If (Prg(PO + 2) = Prg(PO + 2 + J)) And (Prg(PO + 3) = Prg(PO + 3 + J)) And (Prg(PO + 4) = Prg(PO + 4 + J)) And (Prg(PO + 5) = Prg(PO + 5 + J)) Then
    'GoTo Mid2Short
    'End If
    'End Select
    'Next
    'End If

    'End Sub

    'Private Sub CalcLitBits()

    'LitBits = Fix(LitBytes / (MaxLit + 1)) * 9

    'Select Case LitBytes Mod (MaxLit + 1)
    'Case -1
    'LitBits = 0                '0 lit tag bits
    'Case 0
    'LitBits += 2 + 0           '2 lit tag bits
    'Case 1 To 4
    'LitBits += 3 + 2 + 0       '3 lit tag bits + 2 lit sequence length bits
    'Case 5 To 12
    'LitBits += 4 + 3 + 0       '4 lit tag bits + 3 lit sequence length bits
    'Case 13 To MaxLit
    'LitBits += 4 + 5 + 0       '4 lit tag bits + 5 lit sequence length bits
    'End Select

    'End Sub

    Private Function LongMatch() As Boolean
        'On Error GoTo Err

        If DoDebug Then Debug.Print("LongMatch")
        LongMatch = True

        'LONGMATCH

        'If LitCnt > -1 Then CheckShortMidMatch()
        If LitCnt > -1 Then
            'If there is a PostOverLap then we exit function and start FindMatch from next POffset again`
            If Find2ByteMatches() = True Then Exit Function
            'CheckShortMidMatch()
        End If

        CalcNeededBits(True)

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
                GoTo BufferFull
            Else
                If MaxLen > MaxShortLen Then MaxLen = MaxShortLen     'Adjust MaxLen for ShortMatch
                If (MaxOffset <= 64) And (DataFits(1, Bits, LitCnt, CheckIO(POffset - MaxLen + 1))) Then
                    'Add short match here
                    AddLitBits()
                    AddShortM()
                    GoTo BufferFull
                Else
                    'Match does not fit, check if we can add byte as literal
                    CalcNeededBits()
                    If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) Then AddLitByte()        'Add Byte As literal And update LitCnt
                    AddLitBits()
BufferFull:
                    CloseBuff()
                    LongMatch = False    'This will restart FindMatch without advancing POffset
                End If
            End If
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function MidMatch() As Boolean
        'On Error GoTo Err

        If DoDebug Then Debug.Print("MidMatch")
        MidMatch = True

        'MIDMATCH

        'If LitCnt > -1 Then CheckShortMidMatch()
        If LitCnt > -1 Then
            'If there is a PostOverLap then we exit function and start FindMatch from next POffset again`
            If Find2ByteMatches() = True Then Exit Function
            'CheckShortMidMatch()
        End If

        CalcNeededBits(True)

        'Check if Far Match fits
        If DataFits(2, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
            AddLitBits()        'First close last literal sequence
            AddMidM()           'Then add Far Bytes
        Else
            If MaxLen > MaxShortLen Then MaxLen = MaxShortLen
            If (MaxOffset <= 64) And DataFits(1, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
                AddLitBits()
                AddShortM()
                GoTo BufferFull
            Else
                'Match does not fit, check if we can add byte as literal
                CalcNeededBits()

                If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) Then AddLitByte()        'Add Byte As literal And update LitCnt
                AddLitBits()
BufferFull:
                CloseBuff()
                MidMatch = False    'This will restart FindMatch without advancing POffset
            End If
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function ShortMatch() As Boolean
        'On Error GoTo Err

        If DoDebug Then Debug.Print("ShortMatch")
        ShortMatch = True

        'SHORTMATCH

        If LitCnt > -1 Then
            'If there is a PostOverLap then we exit function and start FindMatch from next POffset again`
            If Find2ByteMatches() = True Then Exit Function
            'CheckShortMidMatch()
        End If

        CalcNeededBits(True)

        'We have already reserved space for literal bits, check is match fits
        If DataFits(1, Bits, LitCnt, CheckIO(POffset - MaxLen + 1)) = True Then
            'Yes, we have space
            AddLitBits()    'First close last literal sequence
            AddShortM()       'Then add Near Byte
        Else
            'Match does not fit, check if we can add byte as literal
            CalcNeededBits()

            If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) Then AddLitByte()           'Add byte as literal and update LitCnt

            AddLitBits()
            CloseBuff()
            ShortMatch = False    'This will restart FindMatch without advancing POffset
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

    Private Function Literal() As Boolean
        'On Error GoTo Err

        If DoDebug Then Debug.Print("Literal")
        Literal = True

        'LITERAL

        'CheckShortMidMatch()      'THIS WILL NOT IMPROVE COMPRESSION BUT DEPACKING WILL BE SLOWER...

        CalcNeededBits()

        If DataFits(1, Bits, LitCnt + 1, CheckIO(POffset)) = True Then 'We are adding 1 literal byte and 1+5 bits
            'We have space in the buffer for this byte
            AddLitByte()
        Else    'No space for this literal byte, close sequence
            If LitCnt > -1 Then AddLitBits()

            CloseBuff()
            Literal = False    'Goto Restart
        End If

        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

        Literal = False

    End Function

    Private Sub CalcNeededBits(Optional ForMatch As Boolean = False)

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
                Case 13 To MaxLit
                    Bits += 4 + 5 + 0       '4 lit tag bits + 5 lit sequence length bits
                    'Case MaxLit + 1        'IN THIS VERSION LITCNT CANNOT BE GREATER THAN MAXLIT!!!!
                    'Bits += 4 + 5 + 2       '4 lit tag bits + 5 lit sequence length bits + 2 more lit tag bits
            End Select
        End If

        'If LitCnt >= MaxLit Then MsgBox(Bits.ToString + vbNewLine + LitCnt.ToString)

    End Sub

    Private Function DataFits(ByteLen As Integer, BitLen As Integer, Literals As Integer, Optional SequenceUnderIO As Integer = 0) As Boolean
        'On Error GoTo Err

        If DoDebug Then Debug.Print("DataFits")
        'Do not include closing bits in function call!!!

        Dim NeededBytes As Integer = 1  'Close Byte
        Dim CloseBit As Integer = 0     'CloseBit length - Match selector bit length, not needed most of the time, cannot overlap with Close Byte!!!

        'If (FileUnderIO = True) And (BlockUnderIO = 0) And (SequenceUnderIO = 1) Then
        If (BlockUnderIO = 0) And (SequenceUnderIO = 1) Then

            'Some blocks of the file will go UIO, but sofar this block is not UIO, and next byte is the first one that goes UIO

            NeededBytes += 1    'Need an extra byte if the next byte in sequence is the first one that goes UIO
        End If

        'Check if we have pending Literals
        'If no pending Literals, or Literals=MaxLit, then we need to save 1 bit for match tag
        'Otherwise, next item must be a match, we do not need a match tag
        If (Literals = -1) Or (Literals Mod (MaxLit + 1) = MaxLit) Then CloseBit = 1

        'If Literals >= MaxLit Then MsgBox(BitLen.ToString + vbNewLine + Literals.ToString)

CheckBitLen:
        If BitLen >= 8 Then     'Can be up to 11 bits
            NeededBytes += 1    'Need 1 more byte
            BitLen -= 8
            GoTo CheckBitLen
        End If

        If BitPos < 8 + BitLen + CloseBit Then
            NeededBytes += 1     'Need an extra byte for bit sequence (closing match tag)
        End If

        If ByteCnt >= BitCnt + ByteLen + NeededBytes Then  '>= because NeededBytes also includes Close Byte
            DataFits = True
            'Data will fit
            'If (FileUnderIO = True) And (BlockUnderIO = 0) And (SequenceUnderIO = 1) And (LastByte <> ByteCnt) Then
            If (BlockUnderIO = 0) And (SequenceUnderIO = 1) And (LastByte <> ByteCnt) Then

                'This is the first byte in the block that will go UIO, so lets update the buffer to include the IO flag

                For I As Integer = ByteCnt To AdHiPos            'Move all data to the left in buffer, including AdHi
                    Buffer(I - 1) = Buffer(I)
                Next
                Buffer(AdHiPos) = 0                             'IO Flag to previous AdHi Position
                ByteCnt -= 1                                     'Update ByteCt to next empty position in buffer
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
        'On Error GoTo Err

        If DoDebug Then Debug.Print("AddLongM")

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCnt) = LongMatchTag                   'Long Match Flag = &HF8
        Buffer(ByteCnt - 1) = MaxLen
        Buffer(ByteCnt - 2) = MaxOffset - 1
        ByteCnt -= 3

        POffset -= MaxLen

        LitCnt = -1
        'ResetLit()

        LM += 1

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddMidM()
        'On Error GoTo Err

        If DoDebug Then Debug.Print("AddMidM")

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCnt) = MaxLen * 4                         'Length of match (#$02-#$3f, cannot be #$00 (end byte), and #$01 - distant selector??)
        Buffer(ByteCnt - 1) = MaxOffset - 1
        ByteCnt -= 2

        POffset -= MaxLen

        LitCnt = -1
        'ResetLit()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddShortM()
        'On Error GoTo Err

        If DoDebug Then Debug.Print("AddShortM")

        If (LitCnt = -1) Or (LitCnt Mod (MaxLit + 1) = MaxLit) Then AddRBits(0, 1)   '0		Last Literal Length was -1 or Max, we need the Match Tag

        SaveLastMatch()

        Buffer(ByteCnt) = ((MaxOffset - 1) * 4) + MaxLen
        ByteCnt -= 1

        POffset -= MaxLen

        LitCnt = -1
        ' ResetLit()

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddLitByte()
        'On Error GoTo Err

        If DoDebug Then Debug.Print("AddLitByte")
        'Update bitstream

        'If LitCnt = MaxLit Then 'If LitCnt=Max
        'AddLitBits()        'Then update bitstream
        'LitCnt = -1         'And start new literal sequence
        'OverMaxLit = True   'Next literal will be the first over max
        'End If
        LitCnt += 1             'Increase Literal Counter

        Buffer(ByteCnt) = Prg(POffset)   'Add byte to Byte Stream
        ByteCnt -= 1                     'Update Byte Position Counter

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    'Private Sub CheckShortMidMatch()
    ''On Error GoTo Err
    'Dim MO, ML As Integer

    'Select Case LitCnt Mod (MaxLit + 1)
    'Case 0, 0 + 1 ', 0 + 2, 4 + 1, 4 + 2, 12 + 1, 12 + 2
    ''FindPrePostMatches()
    'End Select
    'CalcOrigNeeds()

    ''Finds 2-byte MidMatches where it would save on bit length
    'Select Case LitCnt Mod (MaxLit + 1)
    'Case 0, 0 + 1
    'If LitCnt <= 1 Then       'LitCnt=45 actually
    ''1 or 2-byte long Literal before next match, check for overlaps here
    'If LastMLen > 1 Then
    'If (CheckOverlap() = False) And (LitCnt <> 0) Then GoTo Second
    'End If
    'Else
    'GoTo Second
    'End If
    'Case 0 + 2, 4 + 1, 4 + 2, 12 + 1, 12 + 2                         'rangemax+1,rangemax+2 (value cannot be 0!!!) (rangemax: -1,0,4,12,44)
    'Second:         For I As Integer = POffset + 2 To MatchStart - 1
    'If (Prg(POffset + 1) = Prg(I)) And (Prg(POffset + 2) = Prg(I + 1)) And (I < POffset + 1 + 256) Then
    'MO = MaxOffset                  'Back up match offset
    'ML = MaxLen                     'Back up match length
    'MaxOffset = I - (POffset + 1)   'Short MidMatch offset
    'MaxLen = 1                      'Short Midmatch length is always 1
    'LitCnt -= 2                     'Literal Count will be 2 less (3 bytes go to whort MidMatch)
    'ByteCnt += 2                    'Go back 2 positions in Byte Stream
    'AddLitBits()                    'Save literal bits
    'AddMidM()                       'Add short MidMatch to Byte Stream
    'POffset += MaxLen               'Restore POffset (AddMidM subtracted MaxLen from POffset)
    'SM2 += 1
    'MaxOffset = MO                  'Restore match offset
    'MaxLen = ML                     'Restore match length
    'Exit Sub
    'End If
    'Next
    'End Select

    'Exit Sub
    'Err:
    'MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    'End Sub

    'Private Function CheckOverlap() As Boolean
    ''On Error GoTo Err

    'CheckOverlap = False

    'Dim PO As Integer = POffset + 1
    'Dim MO, ML As Integer
    ''WE HAVE 1 LITERAL BYTE AFTER A MATCH > 2 BYTES, SO CHECK IF THERE IS AN OVERLAP BACKWARDS
    ''CheckOverlap = True

    'For I As Integer = PO + 1 To MatchStart - 1
    'If I < PO + 256 Then
    ''If current byte= byte(I) and previous byte (last byte of last match)=byte(I+1) then we have an overlap
    'If (Prg(PO) = Prg(I)) And (Prg(PO + 1) = Prg(I + 1)) Then
    'LastMLen -= 1   'Last Match length will be 1 less as last byte is used in new match
    ''Update Last Match
    'If I < PO + 65 Then                        'Overlap is a short match
    'If DataFits(1, 1, -1, CheckIO(POffset)) Then  'Check if new 2-byte match will fit as a short match
    'UpdatePreviousMatch(PO)
    'MO = MaxOffset      'Save next match's offset
    'ML = MaxLen         'and length
    'MaxOffset = I - PO
    'MaxLen = 1
    'LitCnt = -1
    'AddShortM()
    'MaxOffset = MO
    'MaxLen = ML
    'POffset = PO - 1
    'CheckOverlap = True
    'End If
    ''   Else                                            'Overlap is a mid match - THIS WOULD MAKE FILES LONGER!!!!
    ''If DataFits(2, 1, -1, CheckIO(POffset)) Then
    ''UpdatePreviousMatch(PO)
    ''MO = MaxOffset
    ''ML = MaxLen
    ''MaxOffset = I - PO
    ''MaxLen = 1
    ''LitCnt = -1
    ''AddMidM()
    ''MaxOffset = MO
    ''MaxLen = ML
    ''POffset = PO - 1
    ''End If
    'End If
    'LastMLen = 0                                   'Reset last match
    'Exit Function
    'Else
    ''Check Last 2 Lits + 1st Match byte for mid match
    ''Check Last Lit + first 2 Match byts for mid match
    'End If
    'End If
    'Next
    'LastMLen = 0                                   'Reset last match

    'Exit Function
    'Err:
    'MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    'End Function

    'Private Sub UpdatePreviousMatch(Optional PO As Integer = -1)

    'If (LastMOffset <= 64) And (LastMLen <= MaxShortLen) Then   'Update previous match as a short match
    'Buffer(LastByteCt) = ((LastMOffset - 1) * 4) + LastMLen
    ''ByteCnt = LastByteCt - 1 - LitCnt
    'ElseIf LastMLen <= MaxMidLen Then                           'Update previous match as a mid match
    'Buffer(LastByteCt) = LastMLen * 4
    'Buffer(LastByteCt - 1) = LastMOffset - 1
    ''ByteCnt = LastByteCt - 2 - LitCnt
    'Else                                                        'Update previous match as a long match
    'Buffer(LastByteCt) = LongMatchTag
    'Buffer(LastByteCt - 1) = LastMLen
    'Buffer(LastByteCt - 2) = LastMOffset - 1
    ''ByteCnt = LastByteCt - 3 - LitCnt
    'End If

    'Exit Sub

    ''Check if short midmatches could be converted to shortmatches (find new match)

    ''If PO = -1 Then PO = LastPOffset - 2

    'If (LastMLen <= MaxShortLen) And (LastMType = "m") Then
    'For J As Integer = 1 To 63
    'Select Case LastMLen
    'Case 1  'MatchLen=2
    'If (Prg(PO + 2) = Prg(PO + 2 + J)) And (Prg(PO + 3) = Prg(PO + 3 + J)) Then
    ''Debug.Print("M2S 2 bytes" + vbTab + "Part :" + PartCnt.ToString + vbTab + "Block: " + (BlockCnt + 1).ToString)
    'Mid2Short:
    'LastMOffset = J
    'Buffer(LastByteCt) = (LastMOffset * 4) + LastMLen
    ''ByteCnt = LastByteCt - 1
    'For K As Integer = 0 To LitCnt - 1
    'Buffer(LastByteCt - 1 - K) = Buffer(LastByteCt - 2 - K)
    'Next
    ''ByteCnt -= LitCnt
    'ByteCnt += 1
    'Exit For
    'End If
    'Case 2  'MatchLen=3
    'If (Prg(PO + 2) = Prg(PO + 2 + J)) And (Prg(PO + 3) = Prg(PO + 3 + J)) And (Prg(PO + 4) = Prg(PO + 4 + J)) Then
    ''Debug.Print("M2S 3 bytes" + vbTab + "Part :" + PartCnt.ToString + vbTab + "Block: " + (BlockCnt + 1).ToString)
    'GoTo Mid2Short
    'End If
    'Case 3  'MatchLen=4
    'If (Prg(PO + 2) = Prg(PO + 2 + J)) And (Prg(PO + 3) = Prg(PO + 3 + J)) And (Prg(PO + 4) = Prg(PO + 4 + J)) And (Prg(PO + 5) = Prg(PO + 5 + J)) Then
    ''Debug.Print("M2S 4 bytes" + vbTab + "Part :" + PartCnt.ToString + vbTab + "Block: " + (BlockCnt + 1).ToString)
    'GoTo Mid2Short
    'End If
    'End Select
    'Next
    'End If

    'End Sub

    Private Sub AddLitBits()
        'On Error GoTo Err
        If DoDebug Then Debug.Print("AddLitBits")

        If LitCnt = -1 Then Exit Sub

        If LitLenA.Count < LitCnt + 1 Then
            ReDim Preserve LitLenA(LitCnt + 1)
        End If

        LitLenA(LitCnt) += 1

        For I As Integer = 1 To Fix(LitCnt / (MaxLit + 1))
            If FirstOfNext = True Then
                AddRBits(&B11111111, 8)
                FirstOfNext = False
            Else
                AddRBits(&B111111111, 9)
            End If
        Next

        If FirstOfNext = False Then
            AddRBits(1, 1)               'Add Literal Selector if this is not the first (Literal) byte in the buffer
        Else
            FirstOfNext = False
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
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub AddRBits(Bit As Integer, BCnt As Byte)
        'On Error GoTo Err

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
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    'Private Sub AddBits(BitVal As Integer, BitNo As Byte)

    'For I As Integer = BitNo - 1 To 0 Step -1
    'If (BitVal And 2 ^ I) <> 0 Then
    'If BitPos < 8 Then
    'BitPos += 8
    'BitCnt += 1
    'End If
    'Buffer(BitCnt) = Buffer(BitCnt) Or 2 ^ (BitPos - 8)
    'BitPos -= 1
    'End If
    'Next

    'End Sub

    Public Sub CloseBuff()  'CHANGE TO PUBLIC
        'On Error GoTo Err

        If DoDebug Then Debug.Print("CloseBuff")

        Buffer(ByteCnt) = EndTag

        'Utilize unused byte in buffer
        'If ByteCnt - 1 > BitCnt Then
        'If FindMatch() = True Then
        'If (MaxLen <= MaxShortLen) And (MaxOffset <= 64) Then
        'Buffer(ByteCt - 1) = ((MaxOffset - 1) * 4) + MaxLen
        'ByteCt -= 1
        'POffset -= MaxLen
        'Debug.Print("PartCt: " + PartCnt.ToString + vbTab + "BlockCnt: " + (BlockCnt + 1).ToString + vbTab + "BitCt: " + BitCt.ToString + vbTab + "BitPos: " + BitPos.ToString + vbTab + "ByteCt: " + ByteCt.ToString)
        'Debug.Print("PartCt: " + PartCnt.ToString + vbTab + "BlockCnt: " + (BlockCnt + 1).ToString + vbTab + (MaxLen + 1).ToString + " bytes")
        'End If
        'End If
        'End If

        'Close Byte Match Tag = 0, so it is not needed as it is the default value in the next bit position!!!

        Buffer(0) = Buffer(0) And &H7F                    'Delete Compression Bit (Default (i.e. compressed) value is 0)

        'FIND UNCOMPRESSIBLE BLOCKS (only applies to first file in buffer in which case LastByte points at the very first data byte in stream)

        If (MatchStart - POffset <= LastByte) And (MatchStart > LastByte - 1) And (NextFile = False) Then

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
        UpdateByteStream()

        ResetBuffer()               'Resets buffer variables

        NextFile = False            'Reset Next File flag

        If POffset < 0 Then Exit Sub 'We have reached the end of the file -> exit

        'If we have not reached the end of the file, then update buffer

        BlockUnderIO = CheckIO(POffset)          'Check if last byte of prg could go under IO

        Buffer(ByteCnt) = (PrgAdd + POffset) Mod 256
        AdLoPos = ByteCnt

        If BlockUnderIO = 1 Then ByteCnt -= 1

        Buffer(ByteCnt - 1) = Int((PrgAdd + POffset) / 256) Mod 256
        AdHiPos = ByteCnt - 1
        ByteCnt -= 2
        LastByte = ByteCnt               'LastByte = the first byte of the ByteStream after and Address Bytes (253 or 252 with blockCnt)

        Buffer(ByteCnt) = Prg(POffset)   'Copy First Lit Byte to Buffer
        ByteCnt -= 1                     'Update Byte Pos Counter
        MatchStart = POffset            'Update Match Start Flag
        POffset -= 1                    'Next Byte in Prg
        If POffset < 0 Then POffset = 0 'Unless it is <0		NEEDED DO NOT DELETE!!!
        LitCnt = 0                      'Literal counter has 1 value

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub CloseLastBuff()  'CHANGE TO PUBLIC
        'On Error GoTo Err

        If DoDebug Then Debug.Print("CloseLastBuff")

        Buffer(ByteCnt) = EndTag

        'Close Byte Match Tag = 0, so it is not needed as it is the default value in the next bit position!!!

        Buffer(0) = Buffer(0) And &H7F                    'Delete Compression Bit (Default (i.e. compressed) value is 0)

        '--------------------------------------------------------------------------------------------------

        '--------------------------------------------------------------------------------------------------
        If NextFile = False Then
            'Check uncompressed Block IO Status
            'If (CheckIO(MatchStart) Or CheckIO(0) = 1) And (FileUnderIO = True) Then
            If CheckIO(MatchStart) Or CheckIO(0) = 1 Then
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
        'On Error GoTo Err

        ReDim Preserve ByteSt(BlockCnt * 256 - 1)

        For I = 0 To 255
            ByteSt((BlockCnt - 1) * 256 + I) = Buffer(I)
        Next

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Sub SaveLastMatch()
        'On Error GoTo Err

        LastByteCt = ByteCnt
        LastMOffset = MaxOffset
        LastMLen = MaxLen
        LastMType = MaxType
        LastPOffset = POffset
        'LastBytes = MatchBytes
        'LastBits = MatchBits

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Public Sub ResetBuffer() 'CHANGE TO PUBLIC
        'On Error GoTo Err

        ReDim Buffer(255)       'New empty buffer

        'Initialize variables

        If BlockCnt = 0 Then
            ByteCnt = 254        'First file in part, last byte of first block will be BlockCnt
        Else
            ByteCnt = 255        'All other blocks:	Bytestream start at the end of the buffer
        End If

        BitCnt = 0               'Bitstream always starts at the beginning of the buffer

        BitPos = 15             'Reset Bit Position Counter (counts 16 bits backwards: 15-0)

        'DO NOT RESET LitCnt HERE!!! It is needed for match tag check

        Exit Sub
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Sub

    Private Function CheckIO(Offset As Integer) As Integer
        'On Error GoTo Err

        If PrgAdd + Offset < 256 Then       'Are we loading to the Zero Page? If yes, we need to signal it by adding IO Flag
            CheckIO = 1
        Else
            CheckIO = If((PrgAdd + Offset >= &HD000) And (PrgAdd + Offset <= &HDFFF) And (FileUnderIO = True), 1, 0)
        End If


        Exit Function
Err:
        MsgBox(ErrorToString(), vbOKOnly + vbExclamation, Reflection.MethodBase.GetCurrentMethod.Name + " Error")

    End Function

End Module
