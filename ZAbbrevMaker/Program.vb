
'Copyright (C) 2021 Henrik ?sman
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
    Public Const LOW_SCORE_CUTOFF As Integer = 20
    Public Const NUMBER_OF_ABBREVIATIONS As Integer = 96
    Public Const ABBREVIATION_MAX_LENGTH As Integer = 20
    Public Const SPACE_REPLACEMENT As Char = "^"
    Public Const QUOTE_REPLACEMENT As Char = "~"
    Public Const LF_REPLACEMENT As Char = "|"
    Sub Main(args As String())
        Dim forceRoundingTo3 As Boolean = False
        Dim printDebug As Boolean = False
        Dim throwBackLowscorers As Boolean = False
        Dim fastRounding As Boolean = True
        Dim deepRounding As Boolean = True
        Dim fastBeforeDeep As Boolean = False
        Dim gameDirectory As String = Environment.CurrentDirectory
        Dim maxAbbreviationLen As Integer = ABBREVIATION_MAX_LENGTH

        ' Parse arguments
        For i As Integer = 0 To args.Count - 1
            Select Case args(i)
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
                Case "-df"
                    fastRounding = True
                    deepRounding = True
                    fastBeforeDeep = True
                Case "-df"
                    fastRounding = True
                    deepRounding = True
                    fastBeforeDeep = True
                Case "-l"
                    If i < args.Count - 1 AndAlso Integer.TryParse(args(i + 1), maxAbbreviationLen) Then
                        i += 1
                    End If
                Case "-h", "--help", "\?"
                    Console.Error.WriteLine("ZAbbrevMaker 0.2")
                    Console.Error.WriteLine("ZAbbrevMaker [switches] [path-to-game]")
                    Console.Error.WriteLine()
                    Console.Error.WriteLine(" -l nn        Maxlength of abbreviations (default = 20)")
                    Console.Error.WriteLine(" -f           Fast rounding. Try variants (add remove space) to abbreviations")
                    Console.Error.WriteLine("              and see if it gives better savings on account of z-chars rounding.")
                    Console.Error.WriteLine(" -d           Deep rounding. Try up yo 10,000 variants from discarded abbreviations")
                    Console.Error.WriteLine("              and see if it gives better savings on account of z-chars rounding.")
                    Console.Error.WriteLine(" -df          Try deep rounding and then fast rounding, in that order (default).")
                    Console.Error.WriteLine(" -fd          Try fast rounding and then deep rounding, in that order.")
                    Console.Error.WriteLine(" -r3          Always round to 3 for fast and deep rounding. Normally rounding")
                    Console.Error.WriteLine("              to 6 is used for strings stored in high memory for z4+ games.")
                    Console.Error.WriteLine(" -b           Throw all abbreviations that have lower score than last pick back on heap.")
                    Console.Error.WriteLine("              (This only occasionally improves the result, use sparingly.)")
                    Console.Error.WriteLine(" -v           Verbose. Prints progress and intermediate results.")
                    Console.Error.WriteLine(" path-to-game Use this path. If omitted the current path is used.")
                    Console.Error.WriteLine()
                    Console.Error.WriteLine("ZAbbrevMaker executed without any switches in folder with zap-files is")
                    Console.Error.WriteLine("the same as 'ZAbbrevMaker -l 20 -df'.")
                    Exit Sub
                Case Else
                    If IO.Directory.Exists(args(i)) Then
                        gameDirectory = args(i)
                    End If
            End Select
        Next

        SearchForAbbreviations(gameDirectory,
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
    Private Sub SearchForAbbreviations(gameDirectory As String,
                                       maxAbbreviationLen As Integer,
                                       fastRounding As Boolean,
                                       deepRounding As Boolean,
                                       fastBeforeDeep As Boolean,
                                       ForceRoundingTo6 As Boolean,
                                       throwBackLowscorers As Boolean,
                                       printDebug As Boolean)
        Dim searchFor34 As Boolean = False
        Dim searchForCR As Boolean = False
        Dim gameTextList As List(Of gameText) = New List(Of gameText)
        Dim patternDictionary As New Dictionary(Of String, patternData)
        Dim candidateThreshold As Integer = LOW_SCORE_CUTOFF
        Dim lowscoreCutoff As Integer = LOW_SCORE_CUTOFF
        Dim totalSavings As Integer = 0
        Dim zversion As Integer = 3
        Dim gameFilename As String = ""
        Dim packedAddress As Boolean = False

        Dim totalStartTime As Date = Date.Now

        ' Get text from zap-files and store every line in a list of strings.
        ' The ".GSTR", ".STRL", "PRINTI" and "PRINTR" Op-codes contains the text.
        ' Every pattern is stored in a dictionary with the pattern as key.
        Dim startTime As Date = Date.Now
        Console.Error.Write("Indexing files...")
        For Each zapFilename As String In IO.Directory.GetFiles(gameDirectory)
            Dim startPos As Long = 0

            If IO.Path.GetExtension(zapFilename).ToUpper = ".ZAP" And Not zapFilename.Contains("_freq") Then
                If gameFilename = "" OrElse IO.Path.GetFileName(zapFilename).Length < gameFilename.Length Then gameFilename = IO.Path.GetFileName(zapFilename)

                Dim byteText() As Byte = IO.File.ReadAllBytes(zapFilename)

                For i As Long = 5 To byteText.Length - 1
                    Dim opCodeString As String = System.Text.Encoding.ASCII.GetString(byteText, i - 5, 5).ToUpper
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
                            Dim gameTextLine As New gameText(System.Text.Encoding.ASCII.GetString(byteTemp).Replace(String.Concat(QUOTE_REPLACEMENT, QUOTE_REPLACEMENT), QUOTE_REPLACEMENT))
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
                        End If
                    End If
                Next
            End If
        Next
        Console.Error.WriteLine(String.Concat(gameTextList.Count.ToString(), " strings...", Now().AddTicks(-startTime.Ticks).ToString("s.fff \s")))

        ' Add to a Max Heap
        startTime = Date.Now
        Console.Error.Write("Creating max heap...")
        Dim maxHeap As New MoreComplexDataStructures.MaxHeap(Of patternData)
        For Each KPD As KeyValuePair(Of String, patternData) In patternDictionary
            If KPD.Value.Score > LOW_SCORE_CUTOFF Then
                KPD.Value.Savings = KPD.Value.Score
                maxHeap.Insert(KPD.Value)
            End If
        Next
        Console.Error.WriteLine(String.Concat(patternDictionary.Count.ToString(), " patterns...", Now().AddTicks(-startTime.Ticks).ToString("m.ss.fff \s")))

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

        Loop Until currentAbbrev = (NUMBER_OF_ABBREVIATIONS + oversample) Or maxHeap.Count = 0
        Console.Error.WriteLine(String.Concat(Now().AddTicks(-startTime.Ticks).ToString("m.ss.fff \s")))

        If printDebug Then PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, True)

        ' Restore best candidate list to NUMBER_OF_ABBREVIATIONS patterns
        For i As Integer = (NUMBER_OF_ABBREVIATIONS + oversample - 1) To NUMBER_OF_ABBREVIATIONS Step -1
            maxHeap.Insert(bestCandidateList(i))
            bestCandidateList.RemoveAt(i)
        Next

        For pass As Integer = 0 To 1
            Dim prevTotSavings As Integer = 0
            Dim minSavings As Integer = 0
            If (pass = 0 And deepRounding And Not fastBeforeDeep) Or (pass = 1 And deepRounding And fastBeforeDeep) Then
                ' Ok, we now have NUMBER_OF_ABBREVIATIONS abbreviations
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

                If pass = 0 And printDebug Then PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, True)
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
                            If printDebug Then Console.Error.WriteLine("Removing intial space on " & SPACE_REPLACEMENT & FormatAbbreviation(bestCandidateList(i).Key) & ", saving " & deltaSavings.ToString)
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
                            If printDebug Then Console.Error.WriteLine("Removing trailing space on " & FormatAbbreviation(bestCandidateList(i).Key) & SPACE_REPLACEMENT & ", saving " & deltaSavings.ToString)
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
                If pass = 0 And printDebug Then PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, True)
            End If
        Next

        ' Print result
        PrintAbbreviations(AbbrevSort(bestCandidateList, False), gameFilename, False)
        RescoreOptimalParse(gameTextList, bestCandidateList, False, zversion)
        totalSavings = 0
        Dim maxAbbrevs As Integer = bestCandidateList.Count - 1
        If maxAbbrevs >= NUMBER_OF_ABBREVIATIONS Then maxAbbrevs = NUMBER_OF_ABBREVIATIONS - 1
        For i As Integer = 0 To maxAbbrevs
            totalSavings += bestCandidateList(i).Savings
        Next
        Console.Error.WriteLine(String.Concat("Abbrevs would save ", totalSavings.ToString, " z-chars total (~", CInt(totalSavings * 2 / 3).ToString, " bytes)"))
        Console.Error.WriteLine(String.Concat("Totaltid: ", Now().AddTicks(-totalStartTime.Ticks).ToString("m:ss.fff \s")))
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
        Return System.Text.Encoding.ASCII.GetBytes(cChr)(0)
    End Function

    Private Function CalculateCost(sText As String) As Integer
        Dim cCurrent As Char = ""

        Dim iCost As Integer = 0
        For i As Integer = 0 To sText.Length - 1
            cCurrent = CChar(sText.Substring(i, 1))
            If (cCurrent >= "a" And cCurrent <= "z") OrElse cCurrent = SPACE_REPLACEMENT Then
                ' Alphabet A0 and space
                iCost += 1
            ElseIf cCurrent >= "A" And cCurrent <= "Z" Then
                ' Alphabet A1
                iCost += 2
            ElseIf (cCurrent >= "0" And cCurrent <= "9") OrElse ".,!?_#'/\-:()".Contains(cCurrent) OrElse cCurrent = QUOTE_REPLACEMENT OrElse cCurrent = LF_REPLACEMENT Then
                ' Alphabet A2
                iCost += 2
            Else
                ' ZSCII escape, don't count next two characters 
                iCost += 4
                i += 2
            End If
        Next

        Return iCost
    End Function

    Private Function zcharCost(zchar As Char) As Integer
        If (zchar >= "a" And zchar <= "z") OrElse zchar = SPACE_REPLACEMENT Then
            ' Alphabet A0 and space
            Return 1
        ElseIf zchar >= "A" And zchar <= "Z" Then
            ' Alphabet A1
            Return 2
        ElseIf (zchar >= "0" And zchar <= "9") OrElse ".,!?_#'/\-:()".Contains(zchar) OrElse zchar = QUOTE_REPLACEMENT OrElse zchar = LF_REPLACEMENT Then
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
        For i As Integer = 0 To NUMBER_OF_ABBREVIATIONS - 1
            returnList.Add(abbrevList(i).Clone)
        Next

        If sortBottomOfList Then
            Dim tmpList As New List(Of patternData)
            For i As Integer = NUMBER_OF_ABBREVIATIONS To abbrevList.Count - 1
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

            For i As Integer = NUMBER_OF_ABBREVIATIONS To abbrevList.Count - 1
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

        Dim max As Integer = NUMBER_OF_ABBREVIATIONS - 1
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
End Module
