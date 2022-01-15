
'Copyright (C) 2021 Henrik Åsman
'You can redistribute and/or modify this file under the terms of the
'GNU General Public License as published by the Free Software
'Foundation, either version 3 of the License, or (at your option) any
'later version.
'This file is distributed in the hope that it will be useful, but
'WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'General Public License for more details.
'You should have received a copy of the GNU General Public License
'along with this file. If not, see <https://www.gnu.org/licenses/>."
Imports System

Module Program
    Public Const NUMBER_OF_ABBREVIATIONS As Integer = 96
    Public Const ABBREVIATION_MAX_LENGTH As Integer = 20
    Public Const SPACE_REPLACEMENT As Char = "^"
    Public Const QUOTE_REPLACEMENT As Char = "~"
    Public Const LF_REPLACEMENT As Char = "|"

    Private defaultA0 As String = "abcdefghijklmnopqrstuvwxyz"
    Private defaultA1 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Private defaultA2 As String = "0123456789.,!?_#'/\-:()"     ' 3 slots are reserved for an escape char, newline and doublequote
    Private customAlphabet As Boolean = False
    Private A0 As String = defaultA0
    Private A1 As String = defaultA1
    Private A2 As String = defaultA2

    Private Property numberOfAbbrevs As Integer = NUMBER_OF_ABBREVIATIONS

    Private Property textEncoding As System.Text.Encoding = Nothing

    Sub Main(args As String())
        Dim forceRoundingTo3 As Boolean = False
        Dim printDebug As Boolean = False
        Dim throwBackLowscorers As Boolean = False
        Dim fastRounding As Boolean = True
        Dim deepRounding As Boolean = True
        Dim fastBeforeDeep As Boolean = False
        Dim inform6StyleText As Boolean = False
        Dim gameDirectory As String = Environment.CurrentDirectory
        Dim maxAbbreviationLen As Integer = ABBREVIATION_MAX_LENGTH

        ' Parse arguments
        For i As Integer = 0 To args.Count - 1
            Select Case args(i)
                Case "-a"
                    customAlphabet = True
                Case "-a0"
                    If i < args.Count - 1 AndAlso args(i + 1).Length = 26 Then
                        A0 = args(i + 1)
                        i += 1
                    Else
                        Console.Error.WriteLine("ERROR: Can't use defined A0 (needs 26 characters). Using defailt instead.")
                    End If
                Case "-a1"
                    If i < args.Count - 1 AndAlso args(i + 1).Length = 26 Then
                        A1 = args(i + 1)
                        i += 1
                    Else
                        Console.Error.WriteLine("ERROR: Can't use defined A1 (needs 26 characters). Using defailt instead.")
                    End If
                Case "-a2"
                    If i < args.Count - 1 AndAlso args(i + 1).Length = 23 Then
                        A2 = args(i + 1)
                        i += 1
                    Else
                        Console.Error.WriteLine("ERROR: Can't use defined A2 (needs 23 characters). Using defailt instead.")
                    End If
                Case "-r3"
                    forceRoundingTo3 = True
                Case "-v"
                    printDebug = True
                Case "-b"
                    throwBackLowscorers = True
                Case "-f"
                    fastRounding = True
                    deepRounding = False
                Case "-d"
                    fastRounding = False
                    deepRounding = True
                Case "-df"
                    fastRounding = True
                    deepRounding = True
                    fastBeforeDeep = False
                Case "-fd"
                    fastRounding = True
                    deepRounding = True
                    fastBeforeDeep = True
                Case "-l"
                    If i < args.Count - 1 AndAlso Integer.TryParse(args(i + 1), maxAbbreviationLen) Then
                        i += 1
                    End If
                Case "-n"
                    If i < args.Count - 1 AndAlso Integer.TryParse(args(i + 1), numberOfAbbrevs) Then
                        i += 1
                    End If
                Case "-h", "--help", "\?"
                    Console.Error.WriteLine("ZAbbrevMaker 0.8")
                    Console.Error.WriteLine("ZAbbrevMaker [switches] [path-to-game]")
                    Console.Error.WriteLine()
                    Console.Error.WriteLine(" -a           Create a tailor-made alphabet for this game and use it as basis for")
                    Console.Error.WriteLine("              the abbreviations (z5+ only).")
                    Console.Error.WriteLine(" -a0 <string> Define 26 characters for alphabet A0.")
                    Console.Error.WriteLine(" -a1 <string> Define 26 characters for alphabet A1.")
                    Console.Error.WriteLine(" -a2 <string> Define 23 characters for alphabet A2.")
                    Console.Error.WriteLine("              Experimental - works best when text encoding is in ISO-8859-1 (C0 or C1).")
                    Console.Error.WriteLine(" -b           Throw all abbreviations that have lower score than last pick back on heap.")
                    Console.Error.WriteLine("              (This only occasionally improves the result, use sparingly.)")
                    Console.Error.WriteLine(" -c0          Text character set is plain ASCII only.")
                    Console.Error.WriteLine(" -cu          Text character set is UTF-8.")
                    Console.Error.WriteLine(" -c1          Text character set is ISO 8859-1 (Latin1, ANSI).")
                    Console.Error.WriteLine(" -d           Deep rounding. Try up yo 10,000 variants from discarded abbreviations")
                    Console.Error.WriteLine("              and see if it gives better savings on account of z-chars rounding.")
                    Console.Error.WriteLine(" -df          Try deep rounding and then fast rounding, in that order (default).")
                    Console.Error.WriteLine(" -f           Fast rounding. Try variants (add remove space) to abbreviations")
                    Console.Error.WriteLine("              and see if it gives better savings on account of z-chars rounding.")
                    Console.Error.WriteLine(" -fd          Try fast rounding and then deep rounding, in that order.")
                    Console.Error.WriteLine(" -i           Generate output for Inform6. This requires that the file.")
                    Console.Error.WriteLine("              'gametext.txt' is in the gamepath. 'gametext.txt' is generated by:")
                    Console.Error.WriteLine("                 inform6 -r $TRANSCRIPT_FORMAT=1 <game>.inf")
                    Console.Error.WriteLine("              in Inform6 version 6.35 or later. -i always use -r3.")
                    Console.Error.WriteLine(" -l nn        Maxlength of abbreviations (default = 20).")
                    Console.Error.WriteLine(" -n nn        # of abbreviations to generate (default = 96).")
                    Console.Error.WriteLine(" -r3          Always round to 3 for fast and deep rounding. Normally rounding")
                    Console.Error.WriteLine("              to 6 is used for strings stored in high memory for z4+ games.")
                    Console.Error.WriteLine(" -v           Verbose. Prints progress and intermediate results.")
                    Console.Error.WriteLine(" path-to-game Use this path. If omitted the current path is used.")
                    Console.Error.WriteLine()
                    Console.Error.WriteLine("ZAbbrevMaker executed without any switches in folder with zap-files is")
                    Console.Error.WriteLine("the same as 'ZAbbrevMaker -l 20 -df'.")
                    Exit Sub
                Case "-i"
                    inform6StyleText = True
                Case "-c0"
                    textEncoding = System.Text.Encoding.ASCII
                Case "-cu"
                    textEncoding = System.Text.Encoding.UTF8
                Case "-c1"
                    textEncoding = System.Text.Encoding.Latin1
                Case Else
                    If IO.Directory.Exists(args(i)) Then
                        gameDirectory = args(i)
                    End If
            End Select
        Next

        SearchForAbbreviations(gameDirectory,
                               inform6StyleText,
                               maxAbbreviationLen,
                               fastRounding,
                               deepRounding,
                               fastBeforeDeep,
                               forceRoundingTo3,
                               throwBackLowscorers,
                               printDebug)
    End Sub

    'Algorithm, suggested by MTR
    '1.  Calculate abbreviations, score by naive savings
    '2.  Put them in a max-heap.
    '3.  Remove top of heap abbreviation, add To the set of best abbreviations And parse entire corpus.
    '4.  Compute change In savings (Using optimal parse) And declare that to be the new score for the current abbreviation.
    '5.  If the new score is less than the current score for the top-of-heap, remove the current abbreviation from the 
    '    best_abbreviations set And throw it back on the heap.
    '6.  Repeat from step 3 until enough abbreviations are found or the heap is empty.

    ' See:
    ' https://intfiction.org/t/highly-optimized-abbreviations-computed-efficiently/48753
    ' https://gitlab.com/russotto/zilabbrs
    ' https://github.com/hlabrand/retro-scripts
    ' https://www.nuget.org/packages/MoreComplexDataStructures/
    Private Sub SearchForAbbreviations(gameDirectory As String,
                                       inform6StyleText As Boolean,
                                       maxAbbreviationLen As Integer,
                                       fastRounding As Boolean,
                                       deepRounding As Boolean,
                                       fastBeforeDeep As Boolean,
                                       ForceRoundingTo6 As Boolean,
                                       throwBackLowscorers As Boolean,
                                       printDebug As Boolean)
        Try
            Console.Error.WriteLine("ZAbbrevMaker 0.8")

            If Not IO.Directory.Exists(gameDirectory) Then
                Console.Error.WriteLine("ERROR: Can't find specified directory.")
                Exit Sub
            End If

            Dim searchFor34 As Boolean = False
            Dim searchForCR As Boolean = False
            Dim gameTextList As List(Of gameText) = New List(Of gameText)
            Dim patternDictionary As New Dictionary(Of String, patternData)
            Dim totalSavings As Integer = 0
            Dim zversion As Integer = 3
            Dim gameFilename As String = ""
            Dim packedAddress As Boolean = False

            Dim charFreq As New Dictionary(Of Char, Integer)

            Dim totalStartTime As Date = Date.Now

            Dim startTime As Date = Date.Now
            Console.Error.Write("Indexing files...")

            If inform6StyleText Then

                If maxAbbreviationLen > 63 Then
                    Console.Error.WriteLine("WARNING: Max length of abbreviations in Inform are 63 characters. Setting length to 63.")
                    maxAbbreviationLen = 63
                End If

                ' Inform6 text are in "gametext.txt". 
                ' "gametext.txt" is produced by: inform6.exe -r $TRANSCRIPT_FORMAT=1 <gamefile>.inf
                ' Each line is coded
                '   I:info
                '   G:game text
                '   V:veneer text
                '   L:lowmem string
                '   A:abbreviation
                '   D:dict word
                '   O:object name
                '   S:symbol
                '   X:infix
                '   H:game text inline in opcode
                '   W:veneer text inline in opcode
                ' Only text on lines GVLOSHW should be indexed.
                ' ^ means CR and ~ means ".
                ' Candidate strings that contains a @ should not be considered.
                If IO.File.Exists(IO.Path.Combine(gameDirectory, "gametext.txt")) Then

                    If textEncoding Is Nothing Then
                        'Try to autodetect encoding
                        If IsFileUTF8(IO.Path.Combine(gameDirectory, "gametext.txt")) Then
                            textEncoding = System.Text.Encoding.UTF8
                        Else
                            textEncoding = System.Text.Encoding.Latin1
                        End If
                    End If

                    Dim reader As New IO.StreamReader(IO.Path.Combine(gameDirectory, "gametext.txt"), textEncoding)
                    Dim line As String

                    Do
                        line = reader.ReadLine

                        If line IsNot Nothing Then
                            If "GVLOSHW".Contains(line.Substring(0, 1)) Then
                                ' Replace ^, ~ and space
                                line = line.Replace("^", LF_REPLACEMENT)
                                line = line.Replace("~", QUOTE_REPLACEMENT)
                                line = line.Replace(" ", SPACE_REPLACEMENT)

                                ' Add characters to charFreq
                                For i As Integer = 3 To line.Length - 1
                                    Dim c As Char = CChar(line.Substring(i, 1))
                                    If Not (c = LF_REPLACEMENT Or c = QUOTE_REPLACEMENT Or c = SPACE_REPLACEMENT Or ASCII(c) = 27) Then
                                        If Not charFreq.ContainsKey(c) Then charFreq.Add(c, 0)
                                        charFreq(CChar(line.Substring(i, 1))) += 1
                                    End If
                                Next

                                ' Doesn't work yet for Inform6
                                packedAddress = False

                                Dim gameTextLine As New gameText(line.Substring(3))
                                gameTextLine.packedAddress = packedAddress
                                gameTextList.Add(gameTextLine)
                                For Each sKey In ExtractUniquePhrases(gameTextLine.text, 2, maxAbbreviationLen)
                                    If Not sKey.Contains("@") Then ' TODO: Don't ignores all Inform escape-characters
                                        If Not patternDictionary.ContainsKey(sKey) Then
                                            patternDictionary(sKey) = New patternData
                                            patternDictionary(sKey).Cost = CalculateCost(sKey)
                                        End If
                                        patternDictionary(sKey).Frequency += CountOccurrencesReplace(gameTextLine.text, sKey)
                                        patternDictionary(sKey).Key = sKey
                                    End If
                                Next

                            End If
                        End If
                    Loop Until line Is Nothing

                    reader.Close()
                Else
                    Console.Error.WriteLine()
                    Console.Error.WriteLine("ERROR: No 'gametext.txt' in directory.")
                    Exit Sub
                End If
            Else
                ' Get text from zap-files and store every line in a list of strings.
                ' The ".GSTR", ".STRL", "PRINTI" and "PRINTR" Op-codes contains the text.
                ' Every pattern is stored in a dictionary with the pattern as key.
                For Each fileName As String In IO.Directory.GetFiles(gameDirectory)
                    Dim startPos As Long = 0

                    If IO.Path.GetExtension(fileName).ToUpper = ".ZAP" And Not fileName.Contains("_freq") Then
                        If gameFilename = "" OrElse IO.Path.GetFileName(fileName).Length < gameFilename.Length Then gameFilename = IO.Path.GetFileName(fileName)

                        If textEncoding Is Nothing Then
                            If IsFileUTF8(fileName) Then
                                textEncoding = System.Text.Encoding.UTF8
                            Else
                                textEncoding = System.Text.Encoding.Latin1
                            End If

                        End If

                        Dim byteText() As Byte = IO.File.ReadAllBytes(fileName)

                        For i As Long = 5 To byteText.Length - 1
                            Dim opCodeString As String = textEncoding.GetString(byteText, i - 5, 5).ToUpper
                            If opCodeString = ".GSTR" Then
                                searchFor34 = True
                                packedAddress = True
                            End If
                            If opCodeString = ".STRL" Then
                                searchFor34 = True
                                packedAddress = False
                            End If
                            If opCodeString = "RINTI" Then
                                searchFor34 = True
                                packedAddress = False
                            End If
                            If opCodeString = "RINTR" Then
                                searchFor34 = True
                                packedAddress = False
                            End If
                            ' zversion is only relevant if we want to round to 6 zchars for strings in high memory
                            ' ZAPF inserts the file in this order:
                            '   game_freq.zap
                            '   game_data.zap
                            '   game.zap
                            '   game_str.zap
                            '
                            '   dynamic memory : Everything to the label IMPURE (game_data.zap)
                            '   static memory  : Between IMPURE and ENDLOD (game_data.zap)
                            '   high memory    : Everything after ENDLOD
                            '
                            '   Only .GSTR is relevant for rounding to anything other than 3 because text
                            '   in game.zap (inside code) always rounds to 3.
                            If opCodeString = ".NEW " And ForceRoundingTo6 Then zversion = byteText(i) - 48

                            If searchFor34 And byteText(i) = 34 Then
                                startPos = i
                                searchFor34 = False
                                searchForCR = True
                            End If

                            If searchForCR And byteText(i) = 13 Then
                                searchForCR = False

                                ' Replace ", [LF] & Space with printable and legal characters for a Key
                                If (i - startPos - 3) > 0 Then
                                    Dim byteTemp(i - startPos - 3) As Byte
                                    For j As Integer = 0 To byteTemp.Length - 1
                                        Dim byteChar As Byte = byteText(startPos + 1 + j)
                                        If byteChar = 10 Then byteChar = ASCII(LF_REPLACEMENT)
                                        If byteChar = 32 Then byteChar = ASCII(SPACE_REPLACEMENT)
                                        If byteChar = 34 Then byteChar = ASCII(QUOTE_REPLACEMENT)
                                        byteTemp(j) = byteChar
                                    Next

                                    ' Create dictionary. Replace two double-quotes with one (the first is an escape-char). 
                                    Dim gameTextLine As New gameText(textEncoding.GetString(byteTemp).Replace(String.Concat(QUOTE_REPLACEMENT, QUOTE_REPLACEMENT), QUOTE_REPLACEMENT))
                                    gameTextLine.packedAddress = packedAddress
                                    gameTextList.Add(gameTextLine)
                                    For Each sKey In ExtractUniquePhrases(gameTextLine.text, 2, maxAbbreviationLen)
                                        If Not patternDictionary.ContainsKey(sKey) Then
                                            patternDictionary(sKey) = New patternData
                                            patternDictionary(sKey).Cost = CalculateCost(sKey)
                                        End If
                                        patternDictionary(sKey).Frequency += CountOccurrencesReplace(gameTextLine.text, sKey)
                                        patternDictionary(sKey).Key = sKey
                                    Next

                                    ' Add characters to charFreq
                                    For j As Integer = 0 To gameTextLine.text.Length - 1
                                        Dim c As Char = CChar(gameTextLine.text.Substring(j, 1))
                                        If Not (c = LF_REPLACEMENT Or c = QUOTE_REPLACEMENT Or c = SPACE_REPLACEMENT Or ASCII(c) = 27) Then
                                            If Not charFreq.ContainsKey(c) Then charFreq.Add(c, 0)
                                            charFreq(CChar(gameTextLine.text.Substring(j, 1))) += 1
                                        End If
                                    Next

                                End If
                            End If
                        Next
                    End If
                Next
            End If
            Console.Error.WriteLine(String.Concat(gameTextList.Count.ToString(), " strings...", Now().AddTicks(-startTime.Ticks).ToString("s.fff \s")))

            If gameTextList.Count = 0 Then
                Console.Error.WriteLine("ERROR: No data to index.")
                Exit Sub
            End If

            Console.Error.WriteLine(String.Concat("Text encoding: ", textEncoding.BodyName, ", ", textEncoding.EncodingName))

            If customAlphabet Then
                Dim charFreqList As List(Of KeyValuePair(Of Char, Integer)) =
                (From tPair As KeyValuePair(Of Char, Integer) _
                 In charFreq Order By tPair.Value Descending
                 Select tPair).ToList
                Dim alphabet As String = ""
                For i As Integer = 0 To 74
                    alphabet = String.Concat(alphabet, charFreqList(i).Key)
                Next
                A0 = SortAlphabet(alphabet.Substring(0, 26), defaultA0)
                A1 = SortAlphabet(alphabet.Substring(26, 49), defaultA1 & defaultA2)
                A2 = A1.Substring(26)
                A1 = A1.Substring(0, 26)
                If printDebug Then Console.Error.WriteLine("Alphabet = " & Chr(34) & alphabet & Chr(34))
                If inform6StyleText Then PrintAlphabetI6() Else PrintAlphabet()
            End If

            ' Add to a Max Heap
            startTime = Date.Now
            Console.Error.Write("Creating max heap...")
            Dim maxHeap As New MoreComplexDataStructures.MaxHeap(Of patternData)
            For Each KPD As KeyValuePair(Of String, patternData) In patternDictionary

                ' Recalculate cost if where using a newly defined custom alphabet
                If customAlphabet Then KPD.Value.Cost = CalculateCost(KPD.Value.Key)

                KPD.Value.Savings = KPD.Value.Score
                maxHeap.Insert(KPD.Value)
            Next
            Console.Error.WriteLine(String.Concat(patternDictionary.Count.ToString(), " total patterns on heap...", Now().AddTicks(-startTime.Ticks).ToString("m.ss.fff \s")))

            startTime = Date.Now
            Console.Error.Write("Searching for abbreviations with optimal parse...")
            If printDebug Then Console.Error.WriteLine()
            Dim bestCandidateList As New List(Of patternData)
            Dim currentAbbrev As Integer = 0
            Dim previousSavings As Integer = 0
            Dim oversample As Integer = 0
            If throwBackLowscorers Then oversample = 5
            Do
                bestCandidateList.Add(maxHeap.ExtractMax)
                Dim currentSavings As Integer = RescoreOptimalParse(gameTextList, bestCandidateList, False, zversion)
                Dim deltaSavings As Integer = currentSavings - previousSavings
                If deltaSavings < maxHeap.Peek.Savings Then
                    ' If delta savings is less than top of heap then remove current abbrev and reinsert it in heap with new score and try next from heap
                    Dim KPD As patternData = bestCandidateList(currentAbbrev)
                    KPD.Savings = currentSavings - previousSavings
                    bestCandidateList.RemoveAt(currentAbbrev)
                    maxHeap.Insert(KPD)
                Else
                    If printDebug Then Console.Error.WriteLine("Adding abbrev no " & (currentAbbrev + 1).ToString & ": " & FormatAbbreviation(bestCandidateList(currentAbbrev).Key) & ", Total savings: " & currentSavings.ToString)
                    Dim latestSavings As Integer = currentSavings - previousSavings
                    currentAbbrev += 1
                    previousSavings = currentSavings
                    If throwBackLowscorers Then
                        ' put everthing back on heap that has lower savings than latest added
                        Dim bNeedRecalculation As Boolean = False
                        For i As Integer = bestCandidateList.Count - 1 To 0 Step -1
                            If bestCandidateList(i).Savings < latestSavings Then
                                If printDebug Then Console.Error.WriteLine("Removing abbrev: " & FormatAbbreviation(bestCandidateList(i).Key))
                                maxHeap.Insert(bestCandidateList(i))
                                bestCandidateList.RemoveAt(i)
                                i -= 1
                                currentAbbrev -= 1
                                bNeedRecalculation = True
                            End If
                        Next
                        If bNeedRecalculation Then
                            previousSavings = RescoreOptimalParse(gameTextList, bestCandidateList, False, zversion)
                            If printDebug Then Console.Error.WriteLine("Total savings: " & previousSavings.ToString & " - Total Abbrevs: " & currentAbbrev.ToString)
                        End If
                    End If
                End If

            Loop Until currentAbbrev = (numberOfAbbrevs + oversample) Or maxHeap.Count = 0
            Console.Error.WriteLine(String.Concat(Now().AddTicks(-startTime.Ticks).ToString("m.ss.fff \s")))

            If printDebug Then
                If inform6StyleText Then
                    PrintAbbreviationsI6(AbbrevSort(bestCandidateList, False), gameFilename, True)
                Else
                    PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, True)
                End If
            End If

            ' Restore best candidate list to numberOfAbbrevs patterns
            For i As Integer = (numberOfAbbrevs + oversample - 1) To numberOfAbbrevs Step -1
                maxHeap.Insert(bestCandidateList(i))
                bestCandidateList.RemoveAt(i)
            Next

            For pass As Integer = 0 To 1
                Dim prevTotSavings As Integer = 0
                Dim minSavings As Integer = 0
                If (pass = 0 And deepRounding And Not fastBeforeDeep) Or (pass = 1 And deepRounding And fastBeforeDeep) Then
                    ' Ok, we now have numberOfAbbrevs abbreviations
                    ' Recalculate savings taking rounding into account and test a number of candidates to see if they yield a better result
                    startTime = Date.Now
                    Console.Error.Write("Searching for replacements among discarded...")
                    If printDebug Then Console.Error.WriteLine()
                    Dim passes As Integer = 0
                    prevTotSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                    minSavings = prevTotSavings
                    For i As Integer = 0 To bestCandidateList.Count - 1
                        If bestCandidateList(i).Savings < minSavings Then minSavings = bestCandidateList(i).Savings
                    Next
                    Do While passes < 10000 And maxHeap.Count > 0
                        passes += 1
                        Dim runnerup As patternData = maxHeap.ExtractMax
                        'If runnerup.Savings < (minSavings \ 2) Then
                        '    Continue Do
                        'End If
                        Dim replaced As Boolean = False
                        For i = bestCandidateList.Count - 1 To 0 Step -1    ' Search from lowest savings uppward

                            If Not replaced Then
                                If runnerup.Key.StartsWith(bestCandidateList(i).Key) OrElse
                                runnerup.Key.EndsWith(bestCandidateList(i).Key) OrElse
                                bestCandidateList(i).Key.StartsWith(runnerup.Key) OrElse
                                bestCandidateList(i).Key.EndsWith(runnerup.Key) Then
                                    Dim tempCandidate As patternData = bestCandidateList(i)
                                    bestCandidateList.Insert(i, runnerup)
                                    bestCandidateList.RemoveAt(i + 1)
                                    Dim currentSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                                    Dim deltaSavings As Integer = currentSavings - prevTotSavings
                                    If deltaSavings > 0 Then
                                        prevTotSavings = currentSavings
                                        replaced = True
                                        If printDebug Then Console.Error.WriteLine("Replacing " & FormatAbbreviation(tempCandidate.Key) & " with " & FormatAbbreviation(runnerup.Key) & ", saving " & deltaSavings.ToString & ", pass = " & passes.ToString)
                                    Else
                                        bestCandidateList.Insert(i, tempCandidate)
                                        bestCandidateList.RemoveAt(i + 1)
                                    End If
                                End If
                            End If

                        Next
                    Loop
                    Console.Error.WriteLine(String.Concat(Now().AddTicks(-startTime.Ticks).ToString("m.ss.fff \s")))

                    If pass = 0 And printDebug Then
                        If inform6StyleText Then
                            PrintAbbreviationsI6(AbbrevSort(bestCandidateList, False), gameFilename, True)
                        Else
                            PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, True)
                        End If
                    End If
                End If

                If (pass = 1 And fastRounding And Not fastBeforeDeep) Or (pass = 0 And fastRounding And fastBeforeDeep) Then
                    ' Test if we add/remove initial/trailing space
                    startTime = Date.Now
                    Console.Error.Write("Add/remove initial/trailing space...")
                    If printDebug Then Console.Error.WriteLine()
                    prevTotSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                    minSavings = prevTotSavings
                    For i As Integer = 0 To bestCandidateList.Count - 1
                        If bestCandidateList(i).Savings < minSavings Then minSavings = bestCandidateList(i).Savings
                    Next

                    For i As Integer = 0 To bestCandidateList.Count - 1
                        If bestCandidateList(i).Key.StartsWith(SPACE_REPLACEMENT) Then
                            bestCandidateList(i).Key = bestCandidateList(i).Key.Substring(1)
                            bestCandidateList(i).Cost -= 1
                            Dim currentSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                            Dim deltaSavings As Integer = currentSavings - prevTotSavings
                            If deltaSavings > 0 Then
                                ' Keep it
                                prevTotSavings = currentSavings
                                If printDebug Then Console.Error.WriteLine("Removing intial space on " & FormatAbbreviation(SPACE_REPLACEMENT & bestCandidateList(i).Key) & ", saving " & deltaSavings.ToString)
                            Else
                                ' Restore
                                bestCandidateList(i).Key = SPACE_REPLACEMENT & bestCandidateList(i).Key
                                bestCandidateList(i).Cost += 1
                            End If
                        Else
                            bestCandidateList(i).Key = SPACE_REPLACEMENT & bestCandidateList(i).Key
                            bestCandidateList(i).Cost += 1
                            Dim currentSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                            Dim deltaSavings As Integer = currentSavings - prevTotSavings
                            If deltaSavings > 0 Then
                                ' Keep it
                                prevTotSavings = currentSavings
                                If printDebug Then Console.Error.WriteLine("Adding intial space on " & FormatAbbreviation(bestCandidateList(i).Key.Substring(1)) & ", saving " & deltaSavings.ToString)
                            Else
                                ' Restore
                                bestCandidateList(i).Key = bestCandidateList(i).Key.Substring(1)
                                bestCandidateList(i).Cost -= 1
                            End If
                        End If
                    Next

                    For i As Integer = 0 To bestCandidateList.Count - 1
                        If bestCandidateList(i).Key.EndsWith(SPACE_REPLACEMENT) Then
                            bestCandidateList(i).Key = bestCandidateList(i).Key.Substring(0, bestCandidateList(i).Key.Length - 1)
                            bestCandidateList(i).Cost -= 1
                            Dim currentSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                            Dim deltaSavings As Integer = currentSavings - prevTotSavings
                            If deltaSavings > 0 Then
                                ' Keep it
                                prevTotSavings = currentSavings
                                If printDebug Then Console.Error.WriteLine("Removing trailing space on " & FormatAbbreviation(bestCandidateList(i).Key & SPACE_REPLACEMENT) & ", saving " & deltaSavings.ToString)
                            Else
                                ' Restore
                                bestCandidateList(i).Key = bestCandidateList(i).Key & SPACE_REPLACEMENT
                                bestCandidateList(i).Cost += 1
                            End If
                        Else
                            bestCandidateList(i).Key = bestCandidateList(i).Key & SPACE_REPLACEMENT
                            bestCandidateList(i).Cost += 1
                            Dim currentSavings = RescoreOptimalParse(gameTextList, bestCandidateList, True, zversion)
                            Dim deltaSavings As Integer = currentSavings - prevTotSavings
                            If deltaSavings > 0 Then
                                ' Keep it
                                prevTotSavings = currentSavings
                                If printDebug Then Console.Error.WriteLine("Adding trailing space on " & FormatAbbreviation(bestCandidateList(i).Key.Substring(0, bestCandidateList(i).Key.Length - 1)) & ", saving " & deltaSavings.ToString)
                            Else
                                ' Restore
                                bestCandidateList(i).Key = bestCandidateList(i).Key.Substring(0, bestCandidateList(i).Key.Length - 1)
                                bestCandidateList(i).Cost -= 1
                            End If
                        End If
                    Next
                    Console.Error.WriteLine(String.Concat(Now().AddTicks(-startTime.Ticks).ToString("m.ss.fff \s")))
                    If pass = 0 And printDebug Then
                        If inform6StyleText Then
                            PrintAbbreviationsI6(AbbrevSort(bestCandidateList, False), gameFilename, True)
                        Else
                            PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, True)
                        End If
                    End If
                End If
            Next

            ' Print result
            If inform6StyleText Then
                PrintAbbreviationsI6(AbbrevSort(bestCandidateList, False), gameFilename, False)
            Else
                PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, False)
            End If
            RescoreOptimalParse(gameTextList, bestCandidateList, False, zversion)
            totalSavings = 0
            Dim maxAbbrevs As Integer = bestCandidateList.Count - 1
            If maxAbbrevs >= numberOfAbbrevs Then maxAbbrevs = numberOfAbbrevs - 1
            For i As Integer = 0 To maxAbbrevs
                totalSavings += bestCandidateList(i).Savings
            Next
            Console.Error.WriteLine(String.Concat("Abbrevs would save ", totalSavings.ToString, " z-chars total (~", CInt(totalSavings * 2 / 3).ToString, " bytes)"))
            Console.Error.WriteLine(String.Concat("Totaltid: ", Now().AddTicks(-totalStartTime.Ticks).ToString("m:ss.fff \s")))
        Catch ex As Exception
            Console.Error.WriteLine("ERROR: ZAbbrevMaker failed.")
        End Try
    End Sub

    Public Class gameText
        Public Sub New(value As String)
            Me.text = value
        End Sub

        Public text As String = ""
        Private _textSB As Text.StringBuilder
        Public ReadOnly Property textSB As Text.StringBuilder
            Get
                If _textSB Is Nothing Then
                    _textSB = New Text.StringBuilder(Me.text)
                End If
                Return _textSB
            End Get
        End Property

        Public packedAddress As Boolean = False
    End Class
    Public Class patternData
        Implements IComparable(Of patternData)

        Public Key As String = ""
        Public Frequency As Integer = 0
        Public Cost As Integer = 0
        Public Savings As Integer = 0
        Public Tested As Boolean = False

        Public ReadOnly Property Score As Integer
            Get
                ' The total savings for the abbreviation
                ' 2 is the cost in zchars for the link to the abbreviation
                ' The abbreviation also need to be stored once and that
                ' requires the cost rounded up the nearest number 
                ' dividable by 3 (3 zhars per word).
                Return (Frequency * (Cost - 2)) - ((Cost + 2) \ 3) * 3
            End Get
        End Property

        Public Function Clone() As patternData
            Return DirectCast(Me.MemberwiseClone(), patternData)
        End Function

        Public Function CompareTo(other As patternData) As Integer Implements IComparable(Of patternData).CompareTo
            ' Here we compare this instance with other.
            ' If this is supposed to come before other once sorted then
            ' we should return a negative number.
            ' If they are the same, then return 0.
            ' If this one should come after other then return a positive number.
            If Me.Savings > other.Savings Then Return 1
            If Me.Savings < other.Savings Then Return -1
            Return 0
        End Function
    End Class

    Private Function ASCII(cChr As Char) As Integer
        Return textEncoding.GetBytes(cChr)(0)
    End Function

    Private Function CalculateCost(sText As String) As Integer
        Dim cCurrent As Char = ""

        Dim iCost As Integer = 0
        For i As Integer = 0 To sText.Length - 1
            cCurrent = CChar(sText.Substring(i, 1))
            If A0.Contains(cCurrent, StringComparison.Ordinal) OrElse cCurrent = SPACE_REPLACEMENT Then
                ' Alphabet A0 and space
                iCost += 1
            ElseIf A1.Contains(cCurrent, StringComparison.Ordinal) Then
                ' Alphabet A1
                iCost += 2
            ElseIf A2.Contains(cCurrent, StringComparison.Ordinal) OrElse cCurrent = QUOTE_REPLACEMENT OrElse cCurrent = LF_REPLACEMENT Then
                ' Alphabet A2
                iCost += 2
            Else
                ' All other chars cost 4 
                iCost += 4
            End If
        Next

        Return iCost
    End Function

    Private Function zcharCost(zchar As Char) As Integer
        If A0.Contains(zchar, StringComparison.Ordinal) OrElse zchar = SPACE_REPLACEMENT Then
            ' Alphabet A0 and space
            Return 1
        ElseIf A1.Contains(zchar, StringComparison.Ordinal) Then
            ' Alphabet A1
            Return 2
        ElseIf A2.Contains(zchar, StringComparison.Ordinal) OrElse zchar = QUOTE_REPLACEMENT OrElse zchar = LF_REPLACEMENT Then
            ' Alphabet A2
            Return 2
        Else
            Return 4
        End If
    End Function

    Private Function ExtractUniquePhrases(textLine As String, minLen As Integer, maxLen As Integer) As List(Of String)
        Dim patternList As New List(Of String)
        If textLine.Length < maxLen Then maxLen = textLine.Length
        If textLine.Length < minLen Then Return patternList

        For i As Integer = minLen To maxLen
            For j As Integer = 0 To textLine.Length - i
                Dim pattern As String = textLine.Substring(j, i)
                patternList.Add(pattern)
            Next
        Next

        ' Add single characters that have a high charCost (not in A0-A2 alphabet)
        For i As Integer = 0 To textLine.Length - 1
            Dim pattern As String = textLine.Substring(i, 1)
            If zcharCost(pattern) > 3 Then patternList.Add(pattern)
        Next

        Return patternList.Distinct.ToList
    End Function

    Private Function CountOccurrencesReplace(textLine As String, pattern As String) As Integer
        ' Counting with Replace is a little bit faster than counting by Split
        Try
            Return (textLine.Length - textLine.Replace(pattern, String.Empty).Length) \ pattern.Length
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Private Function RescoreOptimalParse(textList As List(Of gameText), abbrevs As List(Of patternData), calculateRoundingPenalty As Boolean, zversion As Integer) As Integer
        ' Parse string using Wagner's optimal parse

        ' Clear frequency from abbrevs
        For Each abbrev As patternData In abbrevs
            abbrev.Frequency = 0
        Next

        Dim roundingPenalty As Integer = 0

        ' Iterate over each string and pick optimal set of abbreviations From abbrevs for this string
        For Each gameTextLine As gameText In textList
            Dim sb As Text.StringBuilder = gameTextLine.textSB
            Dim textLine As String = gameTextLine.text

            Dim abbrLocs = New List(Of List(Of Integer))
            For i As Integer = 0 To sb.Length - 1
                abbrLocs.Add(New List(Of Integer))
            Next

            For i As Integer = 0 To abbrevs.Count - 1
                Dim idx As Integer = textLine.IndexOf(abbrevs(i).Key, StringComparison.Ordinal)
                While idx >= 0
                    abbrLocs(idx).Add(i)
                    idx = textLine.IndexOf(abbrevs(i).Key, idx + 1, StringComparison.Ordinal)
                End While
            Next

            Dim minRemainingCost(sb.Length + 1) As Integer ' Wagner's 'f' or 'F'
            Dim chosenAbbr(sb.Length) As Integer ' -1 for "no abbreviation"
            minRemainingCost(sb.Length) = 0
            Const abbrRefCost As Integer = 2 ' An abbreviation reference Is 2 characters, always.
            For idx As Integer = sb.Length - 1 To 0 Step -1
                Dim charCost As Integer = zcharCost(sb.Chars(idx))
                minRemainingCost(idx) = minRemainingCost(idx + 1) + charCost
                chosenAbbr(idx) = -1
                For Each abbrNo As Integer In abbrLocs(idx)
                    Dim abbrLen As Integer = abbrevs(abbrNo).Key.Length
                    Dim costWithPattern As Integer = abbrRefCost + minRemainingCost(idx + abbrLen)
                    If costWithPattern < minRemainingCost(idx) Then
                        chosenAbbr(idx) = abbrNo
                        minRemainingCost(idx) = costWithPattern
                    End If
                Next
            Next

            'Update frequencies
            For idx As Integer = 0 To sb.Length - 1
                If chosenAbbr(idx) > -1 Then
                    abbrevs(chosenAbbr(idx)).Frequency += 1
                    idx += abbrevs(chosenAbbr(idx)).Key.Length - 1
                End If
            Next

            ' Aggregate rounding penalty for each string.
            ' zchars are 5 bits and are stored in words (16 bits), 3 in each word. 
            ' Depending on rounding 0, 1 or 2 slots can be "wasted" here.
            Dim roundingNumber = 3
            If zversion > 3 AndAlso gameTextLine.packedAddress Then roundingNumber = 6
            If calculateRoundingPenalty Then roundingPenalty -= (roundingNumber - (minRemainingCost(0) Mod roundingNumber)) Mod roundingNumber
        Next

        Dim totalSavings As Integer = 0
        For Each abbrev As patternData In abbrevs
            abbrev.Savings = abbrev.Score
            totalSavings += abbrev.Savings
        Next

        Return totalSavings + roundingPenalty
    End Function

    Private Function AbbrevSort(abbrevList As List(Of patternData), sortBottomOfList As Boolean) As List(Of patternData)
        Dim returnList As New List(Of patternData)
        For i As Integer = 0 To numberOfAbbrevs - 1
            returnList.Add(abbrevList(i).Clone)
        Next

        If sortBottomOfList Then
            Dim tmpList As New List(Of patternData)
            For i As Integer = numberOfAbbrevs To abbrevList.Count - 1
                If abbrevList(i).Score > 0 Then tmpList.Add(abbrevList(i).Clone)
            Next
            tmpList.Sort(Function(firstPair As patternData, nextPair As patternData) CInt(firstPair.Score).CompareTo(CInt(nextPair.Score)))
            tmpList.Reverse()
            For i As Integer = 0 To tmpList.Count - 1
                returnList.Add(tmpList(i).Clone)
            Next
        Else
            returnList.Sort(Function(firstPair As patternData, nextPair As patternData) CInt(firstPair.Key.Length).CompareTo(CInt(nextPair.Key.Length)))
            returnList.Reverse()

            For i As Integer = numberOfAbbrevs To abbrevList.Count - 1
                returnList.Add(abbrevList(i).Clone)
            Next
        End If

        Return returnList
    End Function

    Private Function PrintFormattedAbbreviation(abbreviationNo As Integer, abbreviation As String, frequency As Integer, score As Integer)
        Dim padExtra As Integer = 0
        If abbreviationNo < 10 Then padExtra = 1
        Return String.Concat("        .FSTR FSTR?", abbreviationNo.ToString, ",", FormatAbbreviation(abbreviation).PadRight(20 + padExtra), "; ", frequency.ToString.PadLeft(4), "x, saved ", score)
    End Function

    Private Function FormatAbbreviation(abbreviation As String)
        Return String.Concat(Chr(34), abbreviation.Replace(SPACE_REPLACEMENT, " ").Replace(QUOTE_REPLACEMENT, String.Concat(Chr(34), Chr(34))).Replace(LF_REPLACEMENT, vbLf), Chr(34))
    End Function

    Private Sub PrintAbbreviations(abbrevList As List(Of patternData), gameFilename As String, toError As Boolean)
        Dim outStream As IO.TextWriter = Console.Out
        If toError Then outStream = Console.Error

        outStream.WriteLine("        ; Frequent words file for {0}", gameFilename)
        outStream.WriteLine()

        Dim max As Integer = numberOfAbbrevs - 1
        If abbrevList.Count < max + 1 Then max = abbrevList.Count - 1
        For i As Integer = 0 To max
            outStream.WriteLine(String.Concat(PrintFormattedAbbreviation(i + 1, abbrevList(i).Key, abbrevList(i).Frequency, abbrevList(i).Score)))
        Next

        outStream.WriteLine("WORDS::")

        For i As Integer = 0 To max
            outStream.WriteLine("        FSTR?{0}", i + 1)
        Next

        outStream.WriteLine()
        outStream.WriteLine("        .ENDI")
    End Sub

    Private Sub PrintAbbreviationsI6(abbrevList As List(Of patternData), gameFilename As String, toError As Boolean)
        Dim outStream As IO.TextWriter = Console.Out
        If toError Then outStream = Console.Error

        Dim max As Integer = numberOfAbbrevs - 1
        If abbrevList.Count < max + 1 Then max = abbrevList.Count - 1
        For i As Integer = 0 To max
            Dim line As String = abbrevList(i).Key
            line = line.Replace(SPACE_REPLACEMENT, " ")
            line = line.Replace(LF_REPLACEMENT, "^")
            line = line.Replace("~", QUOTE_REPLACEMENT)
            outStream.WriteLine(String.Concat("Abbreviate ", Chr(34), line, Chr(34), ";"))
        Next
    End Sub

    Private Function IsFileUTF8(fileName As String) As Boolean
        Dim FallbackExp As New System.Text.DecoderExceptionFallback

        Dim fileBytes() As Byte = IO.File.ReadAllBytes(fileName)
        Dim decoderUTF8 = System.Text.Encoding.UTF8.GetDecoder
        decoderUTF8.Fallback = FallbackExp
        Dim IsUTF8 As Boolean = False
        Try
            Dim charCount As Integer = decoderUTF8.GetCharCount(fileBytes, 0, fileBytes.Length)
            IsUTF8 = True
        Catch
            IsUTF8 = False
        End Try

        Return IsUTF8
    End Function

    Private Function SortAlphabet(alphabetIn As String, AlphabetOrg As String) As String
        Dim alphaArray(AlphabetOrg.Length - 1) As Char
        For i As Integer = 0 To alphaArray.Count - 1
            If alphabetIn.Contains(AlphabetOrg.Substring(i, 1), StringComparison.Ordinal) Then
                alphaArray(i) = AlphabetOrg.Substring(i, 1)
            Else
                alphaArray(i) = " "
            End If
        Next
        For i As Integer = 0 To alphaArray.Count - 1
            If Not New String(alphaArray).Contains(alphabetIn.Substring(i, 1), StringComparison.Ordinal) Then
                For j As Integer = 0 To alphaArray.Count - 1
                    If alphaArray(j) = " " Then
                        alphaArray(j) = alphabetIn.Substring(i, 1)
                        Exit For
                    End If
                Next
            End If
        Next
        Return New String(alphaArray)
    End Function

    Private Sub PrintAlphabet()
        Console.Out.WriteLine()
        Console.Out.WriteLine(";" & Chr(34) & "Custom-made alphabet. Insert at beginning of game." & Chr(34))
        Console.Out.WriteLine("<CHRSET 0 " & Chr(34) & A0 & Chr(34) & ">")
        Console.Out.WriteLine("<CHRSET 1 " & Chr(34) & A1 & Chr(34) & ">")
        ' A2, pos 0 - always escape to 10 bit characters
        ' A2, pos 1 - always newline
        ' A2, pos 2 - insert doublequote (as in Inform6)
        Console.Out.WriteLine("<CHRSET 2 " & Chr(34) & "\" & Chr(34) & A2 & Chr(34) & ">")
        Console.Out.WriteLine()
    End Sub

    Private Sub PrintAlphabetI6()
        Console.Out.WriteLine()
        Console.Out.WriteLine("! Custom-made alphabet. Insert at beginning of game.")
        Console.Out.WriteLine("Zcharacter")
        Console.Out.WriteLine("    " & Chr(34) & A0.Replace("@", "@{0040}").Replace("\", "@{005C}") & Chr(34))
        Console.Out.WriteLine("    " & Chr(34) & A1.Replace("@", "@{0040}").Replace("\", "@{005C}") & Chr(34))
        ' A2, pos 0 - always escape to 10 bit characters
        ' A2, pos 1 - always newline
        ' A2, pos 2 - always doublequote
        Console.Out.WriteLine("    " & Chr(34) & A2.Replace("@", "@{0040}").Replace("\", "@{005C}") & Chr(34) & ";")
        Console.Out.WriteLine()
    End Sub
End Module

