Option Explicit

' UTF-8-safe multi-song lyric PowerPoint generator.
' Put this script, lyrics.txt, and the background image(s) in the same folder.
'
' lyrics.txt format (supports MANY songs, one after another):
'
'   [Song Name One]
'   60pt
'   purple
'
'   =V1=
'   ...lyrics...
'   ==
'   ...lyrics...
'   =V2=
'   ...lyrics...
'
'   [Song Name Two]
'   48pt
'   #FF0000
'
'   =V1=
'   ...lyrics...
'
' Rules:
'   - Each song name is wrapped in [ ] and MUST be on its own line.
'   - The line right after the song name may be the font size (e.g. "60pt" or "60").
'   - The line after that may be the font color (named color, #RRGGBB hex, or r,g,b).
'     If either the size or the color is omitted, defaults are used (42pt / black).
'   - Section markers like =V1= =C1= =BRIDGE= start a slide and create a navigation button.
'   - A line containing only == starts a new slide without creating a button on bottom-right.
'   - Background image for a song is named exactly like the song, e.g. "[Song]"
'     -> Song.jpg / Song.jpeg / Song.png. If no exact match exists, the
'     file "background" is used (background.jpg / background.jpeg /
'     background.png). If that does not exist either, a plain white
'     background is used.
'   - [CHANGED] The background is placed ONCE on the SlideMaster
'     (mother slide), not on every slide. To change a song's background
'     afterwards: open the pptx, View -> Slide Master, replace the
'     picture there - every slide updates at once.
'
' Run with: cscript //nologo create-lyrics-ppt.vbs

Const adTypeBinary = 1
Const adTypeText = 2
Const adReadAll = -1
Const msoFalse = 0
Const msoTrue = -1
Const msoTextOrientationHorizontal = 1
Const msoShapeRectangle = 1
Const msoSendToBack = 1
Const ppLayoutBlank = 12
Const ppSaveAsOpenXMLPresentation = 24
Const ppMouseClick = 1
Const ppActionHyperlink = 7
Const ppAlignCenter = 2
Const ppAlignLeft = 1
Const ppVerticalAnchorTop = 1
Const ppVerticalAnchorMiddle = 3
Const ppAutoSizeNone = 0
Const ppAutoSizeShapeToFitText = 1
Const msoAutoSizeNone = 0
Const msoAutoSizeShapeToFitText = 1
Const SLIDE_W = 960
Const SLIDE_H = 540
Const CM_TO_POINTS = 28.3464567
Const LYRIC_BOTTOM_GAP = 65
Const BACKGROUND_TRANSPARENCY = 0.25
Const DEFAULT_FONT_SIZE = 42
Const DEFAULT_COLOR = 0 ' black

Const msoCompressionMixed = 0
Const msoCompressionLow = 1
Const msoCompressionMedium = 2
Const msoCompressionHigh = 3

' [ADDED] Guard/cleanup settings
Const GUARD_MAX_TEXT_BYTES = 5242880
Const GUARD_MIN_TEXT_BYTES = 4
Const GUARD_KILL_ENABLED = True
Const GUARD_ORPHAN_MIN_AGE_MIN = 60

' Module-level state for the current song being built.
Dim fso, scriptFolder, lyricsPath, pptApp
Dim textMargin
Dim g_slideLabels(), g_slideTexts(), g_slideCount
Dim g_buttonLabels(), g_buttonTargets(), g_buttonCount
Dim g_fontSize, g_fontColor
Dim g_compressNote, g_cpicWarned
Dim g_defaultBgSuffix

Main

Sub Main()
    Dim raw, blocks, blockCount, i, block, title
    Dim outputPath, bgPath, createdCount, wasCreated

    textMargin = CM_TO_POINTS

    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    lyricsPath = fso.BuildPath(scriptFolder, "lyrics.txt")
    WScript.Echo ChrW(&H8CC7) & ChrW(&H6599) & ChrW(&H593E) & ChrW(&HFF1A) & " " & scriptFolder

    If Not fso.FileExists(lyricsPath) Then
        MsgBox ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & " lyrics.txt" & ChrW(&H3002) & ChrW(&H4F4D) & ChrW(&H7F6E) & ChrW(&HFF1A) & vbCrLf & scriptFolder, vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & _
            ChrW(&H7C21) & ChrW(&H5831)
        Exit Sub
    End If

    ' [ADDED] Embedded pre-run guard (blocks unsafe runs).
    If Not PreRunGuardLyrics() Then Exit Sub

    ' [CHANGED] Ask which numbered "background-<N>.<ext>" file to use
    ' as the fallback background (see PromptDefaultBackgroundSuffix).
    g_defaultBgSuffix = PromptDefaultBackgroundSuffix(scriptFolder)

    raw = ReadUtf8TextFile(lyricsPath)
    raw = Replace(raw, ChrW(&HFEFF), "")
    raw = Replace(raw, vbCrLf, vbLf)
    raw = Replace(raw, vbCr, vbLf)

    SplitIntoSongs raw, blocks, blockCount
    WScript.Echo ChrW(&H627E) & ChrW(&H5230) & ChrW(&H6B4C) & ChrW(&H66F2) & ChrW(&H6578) & ChrW(&HFF1A) & " " & blockCount
    If blockCount = 0 Then
        MsgBox ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & ChrW(&H6B4C) & ChrW(&H66F2) & ChrW(&H3002) & ChrW(&H8ACB) & ChrW(&H628A) & ChrW(&H6BCF) & ChrW(&H9996) & ChrW(&H6B4C) & ChrW(&H7684) & ChrW(&H6B4C) & ChrW(&H540D) & ChrW(&H7528) & _
            " [ ] " & ChrW(&H5305) & ChrW(&H8D77) & ChrW(&H4F86) & ChrW(&HFF0C) & ChrW(&H4E26) & ChrW(&H5404) & ChrW(&H4F54) & ChrW(&H4E00) & ChrW(&H884C) & ChrW(&H3002), vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & _
            ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
        Exit Sub
    End If

    Set pptApp = Nothing
    wasCreated = False
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set pptApp = CreateObject("PowerPoint.Application")
        wasCreated = True
    End If
    On Error GoTo 0
    If pptApp Is Nothing Then
        MsgBox ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H555F) & ChrW(&H52D5) & " Microsoft PowerPoint" & ChrW(&H3002), vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
        Exit Sub
    End If

    createdCount = 0
    g_compressNote = ""
    g_cpicWarned = False

    For i = 0 To blockCount - 1
        block = blocks(i)
        title = SongTitleFromBlock(block)
        ParseSongBlock block   ' sets g_slideCount, g_slideLabels, g_slideTexts, g_buttonCount, g_buttonLabels, g_buttonTargets, g_fontSize, g_fontColor

        If g_slideCount = 0 Then
            MsgBox ChrW(&H9019) & ChrW(&H9996) & ChrW(&H6B4C) & ChrW(&H6C92) & ChrW(&H6709) & ChrW(&H6BB5) & ChrW(&H843D) & ChrW(&HFF08) & ChrW(&H627E) & ChrW(&H4E0D) & ChrW(&H5230) & " =V1= " & ChrW(&H7B49) & ChrW(&H6A19) & ChrW(&H8A18) & _
                ChrW(&HFF09) & ChrW(&HFF1A) & " " & title, vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
        Else
            bgPath = FindBackgroundForSong(scriptFolder, title)
            outputPath = fso.BuildPath(scriptFolder, SafeFileName(title) & ".pptx")
            If fso.FileExists(outputPath) Then fso.DeleteFile outputPath, True
            CreateSongPptx outputPath, title, bgPath
            createdCount = createdCount + 1
        End If
    Next

    If wasCreated Then
        On Error Resume Next
        pptApp.Quit
        On Error GoTo 0
    End If
    Set pptApp = Nothing
    
    If g_compressNote <> "" Then
        g_compressNote = vbCrLf & vbCrLf & ChrW(&H5716) & ChrW(&H7247) & _
            ChrW(&H58D3) & ChrW(&H7E7C) & ChrW(&HFF1A) & g_compressNote
    End If
    g_compressNote = ""   ' hide g_compressNote on intentionally

    MsgBox ChrW(&H5DF2) & ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&HFF0F) & ChrW(&H53D6) & ChrW(&H4EE3) & " " & createdCount & " " & ChrW(&H500B) & " PowerPoint " & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&H3002) & ChrW(&H4F4D) & ChrW(&H7F6E) & _
        ChrW(&HFF1A) & vbCrLf & scriptFolder & g_compressNote, vbInformation, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)

    ' [ADDED] End-of-run cleanup before Main returns.
    EndOfRunCleanup
End Sub

' ---------------------------------------------------------------
' Split the whole lyrics file into individual [song] blocks.
' ---------------------------------------------------------------
Sub SplitIntoSongs(ByVal raw, ByRef blocks, ByRef blockCount)
    Dim lines, line, i, buf, inBlock
    lines = Split(raw, vbLf)
    blockCount = 0
    inBlock = False
    buf = ""
    For i = 0 To UBound(lines)
        line = Trim(RemoveInvisibleChars(lines(i)))
        If IsSongMarker(line) Then
            If inBlock Then
                If blockCount = 0 Then
                    ReDim blocks(0)
                Else
                    ReDim Preserve blocks(blockCount)
                End If
                blocks(blockCount) = buf
                blockCount = blockCount + 1
            End If
            buf = lines(i)
            inBlock = True
        ElseIf inBlock Then
            buf = buf & vbLf & lines(i)
        End If
    Next
    If inBlock Then
        If blockCount = 0 Then
            ReDim blocks(0)
        Else
            ReDim Preserve blocks(blockCount)
        End If
        blocks(blockCount) = buf
        blockCount = blockCount + 1
    End If
End Sub

' ---------------------------------------------------------------
' Get the song title from a block (the [ ... ] line).
' ---------------------------------------------------------------
Function SongTitleFromBlock(ByVal block)
    Dim lines
    lines = Split(block, vbLf)
    SongTitleFromBlock = Trim(RemoveInvisibleChars(Mid(lines(0), 2, Len(lines(0)) - 2)))
End Function

' ---------------------------------------------------------------
' Parse one song block: font size, font color, then sections and slides.
' Fills module-level g_* variables.
' ---------------------------------------------------------------
Sub ParseSongBlock(ByVal block)
    Dim lines, line, i, stage, currentLabel, currentText, hasSlideForSection
    g_slideCount = 0
    g_buttonCount = 0
    g_fontSize = DEFAULT_FONT_SIZE
    g_fontColor = DEFAULT_COLOR
    lines = Split(block, vbLf)
    stage = 0 ' 0 = expect font size, 1 = expect font color, 2 = lyrics
    currentLabel = ""
    currentText = ""
    hasSlideForSection = False

    For i = 1 To UBound(lines)
        line = Trim(RemoveInvisibleChars(lines(i)))

        If stage = 0 Then
            If line <> "" Then
                If IsSectionMarker(line) Then
                    If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
                        AddSlide currentLabel, currentText
                        hasSlideForSection = True
                    End If
                    currentLabel = NormalizeLabel(line)
                    AddButton currentLabel, g_slideCount
                    currentText = ""
                    hasSlideForSection = False
                    stage = 2
                ElseIf IsSlideBreakMarker(line) Then
                    If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
                        AddSlide currentLabel, currentText
                        hasSlideForSection = True
                    End If
                    currentText = ""
                    stage = 2
                ElseIf IsFontSizeLine(line) Then
                    g_fontSize = ParseFontSize(line)
                    stage = 1
                End If
            End If
        ElseIf stage = 1 Then
            If line <> "" Then
                If IsSectionMarker(line) Then
                    If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
                        AddSlide currentLabel, currentText
                        hasSlideForSection = True
                    End If
                    currentLabel = NormalizeLabel(line)
                    AddButton currentLabel, g_slideCount
                    currentText = ""
                    hasSlideForSection = False
                    stage = 2
                ElseIf IsSlideBreakMarker(line) Then
                    If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
                        AddSlide currentLabel, currentText
                        hasSlideForSection = True
                    End If
                    currentText = ""
                    stage = 2
                Else
                    g_fontColor = ParseColor(line)
                    stage = 2
                End If
            End If
        Else
            ' stage 2: reading lyrics
            If IsSectionMarker(line) Then
                If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
                    AddSlide currentLabel, currentText
                    hasSlideForSection = True
                End If
                currentLabel = NormalizeLabel(line)
                AddButton currentLabel, g_slideCount
                currentText = ""
                hasSlideForSection = False
            ElseIf IsSlideBreakMarker(line) Then
                If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
                    AddSlide currentLabel, currentText
                    hasSlideForSection = True
                End If
                currentText = ""
            ElseIf currentLabel <> "" Then
                ' [ADDED] lyrics-only punctuation -> 2 spaces (see ReplacePunctuation)
                If currentText = "" Then
                    currentText = ReplacePunctuation(lines(i))
                Else
                    currentText = currentText & vbCrLf & ReplacePunctuation(lines(i))
                End If
            End If
        End If
    Next

    If currentLabel <> "" And (currentText <> "" Or Not hasSlideForSection) Then
        AddSlide currentLabel, currentText
    End If
End Sub

' Add one slide to the module-level arrays.
Sub AddSlide(ByVal label, ByVal lyricText)
    If g_slideCount = 0 Then
        ReDim g_slideLabels(0)
        ReDim g_slideTexts(0)
    Else
        ReDim Preserve g_slideLabels(g_slideCount)
        ReDim Preserve g_slideTexts(g_slideCount)
    End If
    g_slideLabels(g_slideCount) = label
    g_slideTexts(g_slideCount) = TrimBlankLines(lyricText)
    g_slideCount = g_slideCount + 1
End Sub

' Add one navigation button targeting a specific slide index.
Sub AddButton(ByVal label, ByVal targetSlideIndex)
    If g_buttonCount = 0 Then
        ReDim g_buttonLabels(0)
        ReDim g_buttonTargets(0)
    Else
        ReDim Preserve g_buttonLabels(g_buttonCount)
        ReDim Preserve g_buttonTargets(g_buttonCount)
    End If
    g_buttonLabels(g_buttonCount) = label
    g_buttonTargets(g_buttonCount) = targetSlideIndex
    g_buttonCount = g_buttonCount + 1
End Sub

' ---------------------------------------------------------------
' File / background helpers
' ---------------------------------------------------------------
Function ReadUtf8TextFile(ByVal filePath)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = adTypeBinary
    stream.Open
    stream.LoadFromFile filePath
    stream.Position = 0
    stream.Type = adTypeText
    stream.Charset = "utf-8"
    ReadUtf8TextFile = stream.ReadText(adReadAll)
    stream.Close
    Set stream = Nothing
End Function

Function FindBackgroundForSong(ByVal folderPath, ByVal songTitle)
    Dim exts, i, p
    exts = Array(".gif", ".jpg", ".jpeg", ".png")

    ' 1) Exact song-name match with the three extensions.
    For i = 0 To UBound(exts)
        p = fso.BuildPath(folderPath, songTitle & exts(i))
        If fso.FileExists(p) Then
            FindBackgroundForSong = p
            Exit Function
        End If
    Next

    ' 2) [CHANGED] Fallback: "background-" + the chosen suffix number
    '    (e.g. "background-2.jpg"), picked once per run by
    '    PromptDefaultBackgroundSuffix. When no "background-<N>" files
    '    exist, or the prompt was left blank/cancelled, g_defaultBgSuffix
    '    is "" and this falls back to the original plain
    '    "background.<ext>" file, unchanged from before.
    '    (Windows file names are case-insensitive, so
    '    Background-2.JPG matches too), checked in the same
    '    extension order.
    For i = 0 To UBound(exts)
        If g_defaultBgSuffix <> "" Then
            p = fso.BuildPath(folderPath, "background-" & g_defaultBgSuffix & exts(i))
        Else
            p = fso.BuildPath(folderPath, "background" & exts(i))
        End If
        If fso.FileExists(p) Then
            FindBackgroundForSong = p
            Exit Function
        End If
    Next

    ' 3) No image at all -> plain white background (empty string).
    FindBackgroundForSong = ""
End Function

' =====================================================================
' [ADDED] Background-suffix picker.
'
' Scans folderPath for image files whose base name starts with
' "background-" (case-insensitive, WITH the hyphen) followed by a
' number, e.g.
'   background-1.jpg  background-2.png  background-10.jpeg
' Only files starting with "background-" are scanned; "background1.jpg"
' (no hyphen) does NOT match, and a plain "background.jpg" (no suffix
' at all) is NOT counted here either - it remains the quiet fallback
' used by FindBackgroundForSong when this function returns "".
'
' If one or more matches are found, an InputBox lists the numbers
' (WITHOUT the hyphen, e.g. "1", "2") and asks the user which one to
' use as the default background for this run. Re-prompts on an
' unrecognized number. Leaving the box blank or pressing Cancel
' returns "" (no numbered default chosen; FindBackgroundForSong then
' falls back to the plain "background.<ext>" file if present, else
' white).
'
' If no "background-<N>" files are found at all, no prompt is shown
' and "" is returned immediately.
' =====================================================================
Function PromptDefaultBackgroundSuffix(ByVal folderPath)
    Dim suffixes, msg, i, chosen, matchFound

    PromptDefaultBackgroundSuffix = ""

    suffixes = ScanNumberedBackgroundSuffixes(folderPath)

    If UBound(suffixes) < 0 Then Exit Function ' none found - keep old behavior

    msg = ChrW(&H5DF2) & ChrW(&H627E) & ChrW(&H5230) & ChrW(&H4EE5) & ChrW(&H4E0B) & _
        " background " & ChrW(&H7DE8) & ChrW(&H865F) & ChrW(&HFF1A) & vbCrLf & vbCrLf
    For i = 0 To UBound(suffixes)
        msg = msg & suffixes(i) & vbCrLf
    Next
    msg = msg & vbCrLf & ChrW(&H8ACB) & ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H8981) & _
        ChrW(&H4F5C) & ChrW(&H70BA) & ChrW(&H9810) & ChrW(&H8A2D) & ChrW(&H80CC) & _
        ChrW(&H666F) & ChrW(&H7684) & " background " & ChrW(&H7DE8) & ChrW(&H865F) & _
        ChrW(&HFF1A)

    Do
        chosen = Trim(InputBox(msg, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & _
            ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831), suffixes(0)))

        If chosen = "" Then
            PromptDefaultBackgroundSuffix = ""
            Exit Function ' blank / cancelled -> old plain "background.<ext>" fallback
        End If

        matchFound = False
        For i = 0 To UBound(suffixes)
            If suffixes(i) = chosen Then
                matchFound = True
                Exit For
            End If
        Next

        If matchFound Then
            PromptDefaultBackgroundSuffix = chosen
            Exit Function
        Else
            MsgBox ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H7684) & ChrW(&H7DE8) & ChrW(&H865F) & _
                ChrW(&H4E0D) & ChrW(&H5B58) & ChrW(&H5728) & ChrW(&H3002) & ChrW(&H8ACB) & _
                ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H8F38) & ChrW(&H5165) & ChrW(&H3002), _
                vbExclamation, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & _
                ChrW(&H7C21) & ChrW(&H5831)
        End If
    Loop
End Function

' [ADDED] Returns a numerically-sorted array of distinct number
' suffixes (WITHOUT the hyphen) found among "background-<N>.<ext>"
' files in folderPath. Only the four supported image extensions are
' scanned (IsImageFile). "background1.jpg" (no hyphen) does NOT match.
Function ScanNumberedBackgroundSuffixes(ByVal folderPath)
    Dim folderObject, fileObject, baseName, suffix, dict, arr, i, j, tmp

    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set folderObject = fso.GetFolder(folderPath)
    If Err.Number <> 0 Or folderObject Is Nothing Then
        Err.Clear
        On Error GoTo 0
        ScanNumberedBackgroundSuffixes = Array()
        Exit Function
    End If
    On Error GoTo 0

    For Each fileObject In folderObject.Files
        If IsImageFile(fileObject.Name) Then
            baseName = fso.GetBaseName(fileObject.Name)
            If Len(baseName) > 11 Then
                If StrComp(Left(baseName, 11), "background-", 1) = 0 Then
                    suffix = Mid(baseName, 12)
                    If IsNumeric(suffix) Then
                        If Not dict.Exists(suffix) Then dict.Add suffix, True
                    End If
                End If
            End If
        End If
    Next

    arr = dict.Keys

    ' Ascending numeric sort (counts are small - plain bubble sort).
    For i = 0 To UBound(arr) - 1
        For j = 0 To UBound(arr) - i - 1
            If CDbl(arr(j)) > CDbl(arr(j + 1)) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next
    Next

    ScanNumberedBackgroundSuffixes = arr
End Function

Function IsImageFile(ByVal name)
    Dim l
    l = LCase(name)
    IsImageFile = (Right(l, 4) = ".gif") Or (Right(l, 4) = ".jpg") Or (Right(l, 5) = ".jpeg") Or (Right(l, 4) = ".png")
End Function

' ---------------------------------------------------------------
' Parsing helpers
' ---------------------------------------------------------------
Function IsSongMarker(ByVal value)
    IsSongMarker = False
    value = Trim(RemoveInvisibleChars(value))

    If Len(value) < 3 Then Exit Function

    Dim isEngBracket, isChiBracket
    isEngBracket = (Left(value, 1) = "[" And Right(value, 1) = "]")
    isChiBracket = (Left(value, 1) = ChrW(&H3010) And Right(value, 1) = ChrW(&H3011))

    If Not (isEngBracket Or isChiBracket) Then Exit Function

    IsSongMarker = True
End Function

Function IsSectionMarker(ByVal value)
    Dim inside
    IsSectionMarker = False
    value = Trim(RemoveInvisibleChars(value))
    If Len(value) < 3 Then Exit Function
    If Left(value, 1) <> "=" Or Right(value, 1) <> "=" Then Exit Function
    inside = Trim(Mid(value, 2, Len(value) - 2))
    If inside = "" Then Exit Function
    If InStr(inside, "=") > 0 Then Exit Function
    IsSectionMarker = True
End Function

Function IsSlideBreakMarker(ByVal value)
    value = Trim(RemoveInvisibleChars(value))
    IsSlideBreakMarker = (value = "==")
End Function

Function NormalizeLabel(ByVal marker)
    NormalizeLabel = UCase(Trim(RemoveInvisibleChars(Mid(marker, 2, Len(marker) - 2))))
End Function

Function IsFontSizeLine(ByVal value)
    Dim clean
    clean = LCase(Replace(Trim(value), " ", ""))
    If Right(clean, 2) = "pt" Then clean = Left(clean, Len(clean) - 2)
    IsFontSizeLine = IsNumeric(clean)
End Function

Function ParseFontSize(ByVal value)
    Dim clean, result
    clean = LCase(Replace(Trim(value), " ", ""))
    If Right(clean, 2) = "pt" Then clean = Left(clean, Len(clean) - 2)
    result = CDbl(clean)
    If result < 8 Then result = 8
    If result > 120 Then result = 120
    ParseFontSize = result
End Function

' ---------------------------------------------------------------
' Font color parsing: named color, #RRGGBB hex, or r,g,b.
' Returns an RGB() value.
' ---------------------------------------------------------------
Function ParseColor(ByVal value)
    Dim v, clean, parts, r, g, b

    v = Trim(value)

    ' #RRGGBB or RRGGBB hex
    clean = LCase(Replace(Replace(v, "#", ""), " ", ""))
    If Len(clean) = 6 Then
        If IsHexDigits(clean) Then
            r = CLng("&H" & Mid(clean, 1, 2))
            g = CLng("&H" & Mid(clean, 3, 2))
            b = CLng("&H" & Mid(clean, 5, 2))
            ParseColor = RGB(r, g, b)
            Exit Function
        End If
    End If

    ' r,g,b
    If InStr(v, ",") > 0 Then
        parts = Split(v, ",")
        If UBound(parts) = 2 And IsNumeric(Trim(parts(0))) And IsNumeric(Trim(parts(1))) And IsNumeric(Trim(parts(2))) Then
            r = Clamp(CInt(Trim(parts(0))), 0, 255)
            g = Clamp(CInt(Trim(parts(1))), 0, 255)
            b = Clamp(CInt(Trim(parts(2))), 0, 255)
            ParseColor = RGB(r, g, b)
            Exit Function
        End If
    End If

    ' named colors
    Select Case LCase(Replace(v, " ", ""))
        Case "black": ParseColor = RGB(0, 0, 0)
        Case "white": ParseColor = RGB(255, 255, 255)
        Case "red": ParseColor = RGB(255, 0, 0)
        Case "darkred": ParseColor = RGB(139, 0, 0)
        Case "crimson": ParseColor = RGB(220, 20, 60)
        Case "green": ParseColor = RGB(0, 128, 0)
        Case "lime": ParseColor = RGB(0, 255, 0)
        Case "darkgreen": ParseColor = RGB(0, 100, 0)
        Case "blue": ParseColor = RGB(0, 0, 255)
        Case "darkblue": ParseColor = RGB(0, 0, 139)
        Case "navy": ParseColor = RGB(0, 0, 128)
        Case "yellow": ParseColor = RGB(255, 255, 0)
        Case "gold": ParseColor = RGB(255, 215, 0)
        Case "orange": ParseColor = RGB(255, 165, 0)
        Case "purple": ParseColor = RGB(128, 0, 128)
        Case "indigo": ParseColor = RGB(75, 0, 130)
        Case "violet": ParseColor = RGB(238, 130, 238)
        Case "magenta": ParseColor = RGB(255, 0, 255)
        Case "pink": ParseColor = RGB(255, 192, 203)
        Case "hotpink": ParseColor = RGB(255, 105, 180)
        Case "cyan": ParseColor = RGB(0, 255, 255)
        Case "teal": ParseColor = RGB(0, 128, 128)
        Case "aqua": ParseColor = RGB(0, 255, 255)
        Case "brown": ParseColor = RGB(165, 42, 42)
        Case "maroon": ParseColor = RGB(128, 0, 0)
        Case "olive": ParseColor = RGB(128, 128, 0)
        Case "gray", "grey": ParseColor = RGB(128, 128, 128)
        Case "silver": ParseColor = RGB(192, 192, 192)
        Case "lightgray", "lightgrey": ParseColor = RGB(211, 211, 211)
        Case "darkgray", "darkgrey": ParseColor = RGB(64, 64, 64)
        Case "coral": ParseColor = RGB(255, 127, 80)
        Case "salmon": ParseColor = RGB(250, 128, 114)
        Case "skyblue": ParseColor = RGB(135, 206, 235)
        Case "turquoise": ParseColor = RGB(64, 224, 208)
        Case Else: ParseColor = DEFAULT_COLOR ' unknown -> black
    End Select
End Function

Function IsHexDigits(ByVal value)
    Dim ch, i
    IsHexDigits = True
    For i = 1 To Len(value)
        ch = Mid(value, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "f")) Then
            IsHexDigits = False
            Exit Function
        End If
    Next
End Function

Function Clamp(ByVal v, ByVal lo, ByVal hi)
    If v < lo Then v = lo
    If v > hi Then v = hi
    Clamp = v
End Function

Function RemoveInvisibleChars(ByVal value)
    value = Replace(value, ChrW(&HFEFF), "")
    value = Replace(value, ChrW(&HA0), " ")
    value = Replace(value, ChrW(&H3000), " ")
    value = Replace(value, ChrW(&H200B), "")
    RemoveInvisibleChars = value
End Function

' =====================================================================
' [ADDED] ReplacePunctuation - lyric lines only.
' Every punctuation mark (full-width CJK + half-width ASCII, quotes,
' brackets, dashes, ellipsis, arrows, dingbats) becomes TWO spaces.
' Called from exactly one place: the stage-2 lyric branch of
' ParseSongBlock. The song title [ ... ], the =V1= / == markers and
' the font-size / font-color lines are NOT touched.
'
' Design notes:
'   * VBScript AscW returns a SIGNED 16-bit Integer, so codepoints
'     above U+7FFF come back negative; c is normalized with +65536.
'   * No giant Array() literal: every punctuation block is its own
'     independent one-line test, so the construct cannot be broken
'     half-way by an editor or copy/paste.
'   * "=" (U+003D) is deliberately NOT replaced (keeps =V1= / ==
'     markers safe even if a lyric line ever contains one).
' =====================================================================
Function ReplacePunctuation(ByVal value)
    Dim i, ch, c, result
    Const REPL = "  "   ' two spaces

    value = "" & value
    result = ""
    For i = 1 To Len(value)
        ch = Mid(value, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536   ' AscW is signed 16-bit
        If IsPunctCode(c) Then
            result = result & REPL
        Else
            result = result & ch
        End If
    Next
    ReplacePunctuation = result
End Function

' [ADDED] helper: True when c (0..65535) is a punctuation codepoint.
' U+3001 IDEOGRAPHIC COMMA, U+FF0C FULLWIDTH COMMA etc. all land here.
Function IsPunctCode(ByVal c)
    IsPunctCode = True
    ' ASCII  ! " # $ % & ' ( ) * + , - . /      ( = is NOT included )
    If (c >= &H21 And c <= &H2F) Then Exit Function
    ' ASCII  : ; <      and      > ? @
    If (c >= &H3A And c <= &H3C) Then Exit Function
    If (c >= &H3E And c <= &H40) Then Exit Function
    ' ASCII  [ \ ] ^ _ `      and      { | } ~
    If (c >= &H5B And c <= &H60) Then Exit Function
    If (c >= &H7B And c <= &H7E) Then Exit Function
    ' Latin-1 punctuation  A1 AB-AE B0-B1 B6-B7 BB BF
    If c = &HA1 Then Exit Function
    If (c >= &HAB And c <= &HAE) Then Exit Function
    If (c >= &HB0 And c <= &HB1) Then Exit Function
    If (c >= &HB6 And c <= &HB7) Then Exit Function
    If c = &HBB Then Exit Function
    If c = &HBF Then Exit Function
    ' General Punctuation: dashes, quotes, bullets, ellipsis ...
    If (c >= &H2010 And c <= &H2027) Then Exit Function
    If (c >= &H2030 And c <= &H205E) Then Exit Function
    ' super/subscript parens, numero, arrows, minus sign
    If (c >= &H207D And c <= &H207E) Then Exit Function
    If (c >= &H208D And c <= &H208E) Then Exit Function
    If c = &H2116 Then Exit Function
    If (c >= &H2190 And c <= &H21FF) Then Exit Function
    If c = &H2212 Then Exit Function
    ' technical angle / ceiling / floor brackets, box drawing
    If (c >= &H2308 And c <= &H230B) Then Exit Function
    If (c >= &H2329 And c <= &H232A) Then Exit Function
    If (c >= &H2500 And c <= &H2501) Then Exit Function
    ' dingbats + ornamental / math brackets
    If (c >= &H2701 And c <= &H2775) Then Exit Function
    If (c >= &H27E6 And c <= &H27EB) Then Exit Function
    If (c >= &H2983 And c <= &H2998) Then Exit Function
    If (c >= &H29D8 And c <= &H29DB) Then Exit Function
    If (c >= &H29FC And c <= &H29FD) Then Exit Function
    ' CJK punctuation  U+3001-U+3003 , U+3008-U+301F , U+3030
    If (c >= &H3001 And c <= &H3003) Then Exit Function
    If (c >= &H3008 And c <= &H301F) Then Exit Function
    If c = &H3030 Then Exit Function
    ' vertical forms + CJK compatibility + small forms
    If (c >= &HFE10& And c <= &HFE19&) Then Exit Function
    If (c >= &HFE30& And c <= &HFE6B&) Then Exit Function
    ' fullwidth forms  FF01-FF0F  FF1A-FF20  FF3B-FF40  FF5B-FF65
    If (c >= &HFF01& And c <= &HFF0F&) Then Exit Function
    If (c >= &HFF1A& And c <= &HFF20&) Then Exit Function
    If (c >= &HFF3B& And c <= &HFF40&) Then Exit Function
    If (c >= &HFF5B& And c <= &HFF65&) Then Exit Function
    If (c >= &HFF9E& And c <= &HFF9F&) Then Exit Function
    If (c >= &HFFE0& And c <= &HFFE6&) Then Exit Function
    If (c >= &HFFE8& And c <= &HFFEE&) Then Exit Function
    IsPunctCode = False
End Function

Function TrimBlankLines(ByVal value)
    Dim parts, firstIndex, lastIndex, i, result
    value = Replace(value, vbCrLf, vbLf)
    value = Replace(value, vbCr, vbLf)
    parts = Split(value, vbLf)
    firstIndex = 0
    lastIndex = UBound(parts)

    Do While firstIndex <= lastIndex And Trim(parts(firstIndex)) = ""
        firstIndex = firstIndex + 1
    Loop
    Do While lastIndex >= firstIndex And Trim(parts(lastIndex)) = ""
        lastIndex = lastIndex - 1
    Loop

    result = ""
    For i = firstIndex To lastIndex
        If result = "" Then
            result = parts(i)
        Else
            result = result & vbCrLf & parts(i)
        End If
    Next
    TrimBlankLines = result
End Function

' ---------------------------------------------------------------
' Build a single song's presentation (one file per song).
' Uses the module-level g_* arrays for this song's sections.
' ---------------------------------------------------------------
Sub CreateSongPptx(ByVal outputPath, ByVal title, ByVal bgPath)
    Dim pptPres, slideRefs, btnSlideRefs, i
    Dim tempJpg, stepMsg, bgUse

    Set pptPres = pptApp.Presentations.Add(msoFalse)
    pptPres.PageSetup.SlideWidth = SLIDE_W
    pptPres.PageSetup.SlideHeight = SLIDE_H

    bgUse = bgPath
    If bgPath <> "" Then
        If IsGifFile(bgPath) Then
            ' Animated GIF: WIA only reads frame 1 and its output is
            ' JPEG (static), so pre-scaling would kill the animation.
            ' Embed the original file untouched - PowerPoint plays the
            ' GIF in the slide show, and identical images on multiple
            ' slides are stored only once.
            g_compressNote = g_compressNote & ChrW(&H3010) & title & ChrW(&H3011) & _
                ChrW(&H80CC) & ChrW(&H666F) & ChrW(&H5716) & " GIF " & _
                ChrW(&H4FDD) & ChrW(&H7559) & ChrW(&H539F) & ChrW(&H5716) & _
                ChrW(&HFF08) & ChrW(&H52D5) & ChrW(&H756B) & ChrW(&H4FDD) & _
                ChrW(&H7559) & ChrW(&HFF0C) & ChrW(&H672A) & ChrW(&H9810) & _
                ChrW(&H7E7C) & ChrW(&HFF09) & ChrW(&H3002) & vbCrLf
        Else
            tempJpg = scriptFolder & "\~bg-150dpi-tmp.jpg"
            If fso.FileExists(tempJpg) Then fso.DeleteFile tempJpg, True
            If Background150Dpi(bgPath, tempJpg, stepMsg) And fso.FileExists(tempJpg) Then
                bgUse = tempJpg
                g_compressNote = g_compressNote & ChrW(&H3010) & title & ChrW(&H3011) & _
                    ChrW(&H80CC) & ChrW(&H666F) & ChrW(&H5716) & ChrW(&H5DF2) & ChrW(&H9810) & _
                    ChrW(&H7E7C) & " 150 ppi" & ChrW(&H3002) & vbCrLf
            Else
                If fso.FileExists(tempJpg) Then fso.DeleteFile tempJpg, True
                g_compressNote = g_compressNote & ChrW(&H3010) & title & ChrW(&H3011) & _
                    "WIA " & ChrW(&H7E7C) & ChrW(&H5716) & ChrW(&H5931) & ChrW(&H6557) & _
                    ChrW(&HFF1A) & stepMsg & ChrW(&HFF0C) & ChrW(&H5DF2) & _
                    ChrW(&H6539) & ChrW(&H7528) & ChrW(&H539F) & ChrW(&H5716) & _
                    ChrW(&H3002) & vbCrLf
            End If
        End If
    End If

    ' [CHANGED] background image + fade overlay now live on the
    ' SlideMaster (mother slide): swap it once in Slide Master view
    ' and every slide of this song updates at once.
    SetupMasterBackground pptPres, bgUse

    ReDim slideRefs(g_slideCount - 1)
    For i = 0 To g_slideCount - 1
        Set slideRefs(i) = pptPres.Slides.Add(i + 1, ppLayoutBlank)
        CreateLyricSlide slideRefs(i), title, g_slideLabels(i), g_slideTexts(i)
    Next

    If g_buttonCount > 0 Then
        ReDim btnSlideRefs(g_buttonCount - 1)
        For i = 0 To g_buttonCount - 1
            Set btnSlideRefs(i) = slideRefs(g_buttonTargets(i))
        Next

        For i = 0 To g_slideCount - 1
            AddNavigationButtons slideRefs(i), g_buttonLabels, btnSlideRefs, g_buttonCount, g_slideLabels(i)
        Next
    End If

    CompressPptx pptPres, bgUse

    pptPres.SaveAs outputPath, ppSaveAsOpenXMLPresentation
    pptPres.Close
    On Error Resume Next
    If bgUse <> bgPath Then fso.DeleteFile bgUse, True
    Err.Clear
    On Error GoTo 0
    Set pptPres = Nothing
End Sub

' ---------------------------------------------------------------------
' [ADDED] SetupMasterBackground - puts the song background on the
' presentation's SlideMaster (the "mother slide") instead of copying
' the picture onto every slide. The user can afterwards open
' View -> Slide Master, replace the picture once, and ALL slides of
' the song show the new background at once.
' The 25% white fade overlay is placed on the master too, directly
' above the picture, so the faded look is preserved after a manual
' swap. With no image file at all, a solid white master background
' is used (same look as before).
' NOTE: master shapes always render BEHIND slide-level shapes, so the
' per-slide lyric text boxes, footers and navigation buttons stay on
' top - exactly like before.
' NOTE (GIF, per user decision): an animated GIF is embedded on the
' master untouched; PowerPoint may show it as a still image in the
' slide show.
' ---------------------------------------------------------------------
Sub SetupMasterBackground(ByVal pres, ByVal bgPath)
    Dim msh, bg, whiteOverlay
    Set msh = pres.SlideMaster.Shapes

    If bgPath <> "" Then
        ' Full-master picture at the bottom of the master stacking order.
        Set bg = msh.AddPicture(bgPath, msoFalse, msoTrue, 0, 0, SLIDE_W, SLIDE_H)
        bg.ZOrder msoSendToBack

        ' A translucent white overlay above the picture fades it.
        Set whiteOverlay = msh.AddShape(msoShapeRectangle, 0, 0, SLIDE_W, SLIDE_H)
        whiteOverlay.Fill.ForeColor.RGB = RGB(255, 255, 255)
        whiteOverlay.Fill.Transparency = BACKGROUND_TRANSPARENCY
        whiteOverlay.Line.Visible = msoFalse
    Else
        ' No image at all -> solid white master background.
        Set whiteOverlay = msh.AddShape(msoShapeRectangle, 0, 0, SLIDE_W, SLIDE_H)
        whiteOverlay.Fill.ForeColor.RGB = RGB(255, 255, 255)
        whiteOverlay.Fill.Solid
        whiteOverlay.Line.Visible = msoFalse
    End If
End Sub

' ---------------------------------------------------------------------
' Background150Dpi - writes a scaled JPEG copy of srcPath to destJpg
' using Windows' built-in WIA (no PowerPoint API involved).
' Slide is SLIDE_W x SLIDE_H points; 150 ppi at that display size is
' maxW x maxH pixels. WIA only scales down, never up. The original
' file is never modified. On any failure returns False and puts
' "step (err code description)" into stepMsg - caller falls back
' to the original image.
' ---------------------------------------------------------------------
Function Background150Dpi(ByVal srcPath, ByVal destJpg, ByRef stepMsg)
    Dim img, proc, maxW, maxH, step
    Background150Dpi = False
    stepMsg = ""
    maxW = CLng(SLIDE_W * 150 / 72)
    maxH = CLng(SLIDE_H * 150 / 72)
    On Error Resume Next
    step = "WIA.ImageFile"
    Set img = CreateObject("WIA.ImageFile")
    step = "LoadFile"
    img.LoadFile srcPath
    step = "WIA.ImageProcess"
    Set proc = CreateObject("WIA.ImageProcess")
    step = "Scale filter"
    proc.Filters.Add proc.FilterInfos("Scale").FilterID
    step = "Scale props"
    proc.Filters(1).Properties("MaximumWidth") = maxW
    proc.Filters(1).Properties("MaximumHeight") = maxH
    proc.Filters(1).Properties("PreserveAspectRatio") = True
    step = "Convert filter"
    proc.Filters.Add proc.FilterInfos("Convert").FilterID
    step = "Convert props"
    proc.Filters(2).Properties("FormatID") = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    proc.Filters(2).Properties("Quality") = 85
    step = "Apply"
    Set img = proc.Apply(img)
    step = "SaveFile"
    img.SaveFile destJpg
    If Err.Number <> 0 Then
        stepMsg = step & " (err " & Err.Number & " " & Err.Description & ")"
    Else
        Background150Dpi = True
    End If
    Err.Clear
    On Error GoTo 0
    Set img = Nothing
    Set proc = Nothing
End Function

' ---------------------------------------------------------------------
' IsGifFile - detects a GIF by its magic bytes (GIF87a/GIF89a), NOT by
' file extension (a GIF named .jpg is still a GIF). Uses the built-in
' ADODB.Stream. On any failure returns False (normal path continues -
' never breaks).
' 要保留 GIF 動畫 => PowerPoint → 檔案 → 選項 → 進階 → 影像大小與品質 → 勾選「Do not compress images in file」 → 確定 → 重跑腳本
' ---------------------------------------------------------------------
Function IsGifFile(ByVal filePath)
    IsGifFile = False
    On Error Resume Next
    Dim ad
    Set ad = CreateObject("ADODB.Stream")
    If Err.Number = 0 Then
        ad.Type = 1
        ad.Open
        ad.LoadFromFile filePath
        IsGifFile = (Left(ad.Read(4), 4) = "GIF8")
        ad.Close
    End If
    Err.Clear
    On Error GoTo 0
    Set ad = Nothing
End Function

' ---------------------------------------------------------------------
' CompressPptx - best-effort second pass with PowerPoint's own
' CompressPictures (must run BEFORE SaveAs). Some builds do not
' expose it (late binding -> 438); that is fine, the WIA pre-scale
' above already did the real work. Reported only once, never breaks
' the run.
' ---------------------------------------------------------------------
Sub CompressPptx(ByVal pres, ByVal bgFile)
    If IsGifFile(bgFile) Then Exit Sub
    On Error Resume Next
    pres.CompressPictures msoCompressionHigh, 150
    If Err.Number <> 0 Then
        If Not g_cpicWarned Then
            g_cpicWarned = True
            g_compressNote = g_compressNote & "PPT " & ChrW(&H5167) & ChrW(&H5EFA) & _
                ChrW(&H4E0D) & ChrW(&H652F) & ChrW(&H6350) & ChrW(&HFF08) & _
                ChrW(&H932F) & ChrW(&H8AA4) & " 438" & ChrW(&HFF09) & ChrW(&HFF0C) & _
                ChrW(&H5DF2) & ChrW(&H81EA) & ChrW(&H52D5) & ChrW(&H6539) & _
                ChrW(&H7528) & " WIA " & ChrW(&H9810) & ChrW(&H7E7C) & ChrW(&H5716) & _
                ChrW(&HFF0C) & ChrW(&H7D50) & ChrW(&H679C) & ChrW(&H4E0D) & _
                ChrW(&H53D7) & ChrW(&H5F71) & ChrW(&H97FF) & ChrW(&H3002) & vbCrLf
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' [CHANGED] background + fade overlay moved to the SlideMaster (see
' SetupMasterBackground); this sub now only builds the per-slide text.
Sub CreateLyricSlide(ByVal sld, ByVal title, ByVal label, ByVal lyricText)
    Dim lyricBox, footer, footerText

    Set lyricBox = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, textMargin, textMargin, SLIDE_W - textMargin * 2, SLIDE_H - textMargin - LYRIC_BOTTOM_GAP)
    lyricBox.TextFrame.AutoSize = ppAutoSizeNone
    lyricBox.TextFrame2.AutoSize = msoAutoSizeNone
    lyricBox.TextFrame.VerticalAnchor = ppVerticalAnchorTop
    lyricBox.TextFrame.MarginLeft = 0
    lyricBox.TextFrame.MarginRight = 0
    lyricBox.TextFrame.MarginTop = 0
    lyricBox.TextFrame.MarginBottom = 0
    lyricBox.TextFrame.WordWrap = msoTrue
    lyricBox.TextFrame.TextRange.Text = lyricText

    With lyricBox.TextFrame.TextRange.Font
        On Error Resume Next
        .NameFarEast = "Microsoft JhengHei"
        Err.Clear
        On Error GoTo 0
        .Name = "Microsoft JhengHei"
        .Size = g_fontSize
        .Bold = msoTrue
        .Color.RGB = g_fontColor
    End With
    lyricBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft

    Set footer = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, textMargin, SLIDE_H - 30, 360, 24)
    footer.TextFrame.AutoSize = ppAutoSizeNone
    footer.TextFrame.VerticalAnchor = ppVerticalAnchorTop
    footer.TextFrame.MarginLeft = 0
    footer.TextFrame.MarginRight = 0
    footer.TextFrame.MarginTop = 0
    footer.TextFrame.MarginBottom = 0
    If Trim(label) <> "" Then
        footerText = title & " " & LCase(label)
    Else
        footerText = title
    End If
    footer.TextFrame.TextRange.Text = footerText
    With footer.TextFrame.TextRange.Font
        On Error Resume Next
        .NameFarEast = "Microsoft JhengHei"
        Err.Clear
        On Error GoTo 0
        .Name = "Microsoft JhengHei"
        .Size = 18
        .Bold = msoTrue
        .Color.RGB = g_fontColor
    End With
    footer.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignLeft
End Sub

Sub AddNavigationButtons(ByVal sld, ByRef labels, ByRef targets, ByVal count, ByVal activeLabel)
    Dim i, buttonHeight, gap, padH, minButtonWidth, totalWidth, startX, buttonY, currentX
    Dim btns(), btn, targetSubAddress

    If count <= 0 Then Exit Sub

    buttonHeight = 28
    gap = 3
    padH = 10
    minButtonWidth = 42
    buttonY = SLIDE_H - 40

    ReDim btns(count - 1)
    totalWidth = 0

    ' 1) Create all buttons and auto-fit their widths without wrapping.
    For i = 0 To count - 1
        Set btn = sld.Shapes.AddShape(msoShapeRectangle, 0, buttonY, minButtonWidth, buttonHeight)
        btn.Line.ForeColor.RGB = RGB(100, 100, 100)
        If UCase(labels(i)) = UCase(activeLabel) Then
            btn.Fill.ForeColor.RGB = RGB(255, 212, 92)
        Else
            btn.Fill.ForeColor.RGB = RGB(230, 235, 242)
        End If

        btn.TextFrame.WordWrap = msoFalse
        btn.TextFrame.AutoSize = ppAutoSizeShapeToFitText
        btn.TextFrame.MarginLeft = padH
        btn.TextFrame.MarginRight = padH
        btn.TextFrame.MarginTop = 0
        btn.TextFrame.MarginBottom = 0
        btn.TextFrame.VerticalAnchor = ppVerticalAnchorMiddle

        btn.TextFrame.TextRange.Text = labels(i)
        With btn.TextFrame.TextRange.Font
            .Name = "Arial"
            .Size = 15
            .Bold = msoTrue
            .Color.RGB = RGB(0, 0, 0)
        End With
        btn.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter

        ' Lock size with minimum width check and fixed height/Y
        btn.TextFrame.AutoSize = ppAutoSizeNone
        If btn.Width < minButtonWidth Then btn.Width = minButtonWidth
        btn.Height = buttonHeight
        btn.Top = buttonY

        targetSubAddress = CStr(targets(i).SlideID) & "," & CStr(targets(i).SlideIndex) & ","
        With btn.ActionSettings(ppMouseClick)
            .Action = ppActionHyperlink
            .Hyperlink.Address = ""
            .Hyperlink.SubAddress = targetSubAddress
        End With

        Set btns(i) = btn
        totalWidth = totalWidth + btn.Width
    Next

    ' 2) Calculate startX so buttons align to bottom-right of slide.
    totalWidth = totalWidth + (count - 1) * gap
    startX = SLIDE_W - totalWidth - 28

    ' 3) Position buttons horizontally from left to right.
    currentX = startX
    For i = 0 To count - 1
        btns(i).Left = currentX
        currentX = currentX + btns(i).Width + gap
    Next
End Sub

Function SafeFileName(ByVal value)
    Dim invalidChars, i
    invalidChars = Array("\\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = 0 To UBound(invalidChars)
        value = Replace(value, invalidChars(i), "_")
    Next
    value = Trim(value)
    If value = "" Then value = "Lyrics"
    SafeFileName = value
End Function


' =====================================================================
' [ADDED] Embedded pre-run guard for the lyrics generator.
' Returns False (and explains why) when the run must not proceed.
' =====================================================================

Function PreRunGuardLyrics()
    Dim app, isVisible, presCount, i, unsavedList, notes, sz

    PreRunGuardLyrics = False
    notes = ""
    unsavedList = ""
    Set app = Nothing

    On Error Resume Next
    Set app = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Or app Is Nothing Then
        Err.Clear
        On Error GoTo 0
    Else
        isVisible = app.Visible
        If Err.Number <> 0 Then
            Err.Clear
            Set app = Nothing
            On Error GoTo 0
            MsgBox "PowerPoint " & ChrW(&H6B63) & ChrW(&H5728) & ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H4F46) & ChrW(&H6C92) & ChrW(&H6709) & ChrW(&H56DE) & ChrW(&H61C9) & ChrW(&HFF08) & ChrW(&H53EF) & ChrW(&H80FD) & ChrW(&H5F48) & _
                ChrW(&H51FA) & ChrW(&H5C0D) & ChrW(&H8A71) & ChrW(&H65B9) & ChrW(&H584A) & ChrW(&HFF09) & ChrW(&H3002) & _
                   ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H624B) & ChrW(&H52D5) & ChrW(&H8655) & ChrW(&H7406) & ChrW(&HFF0C) & ChrW(&H7136) & ChrW(&H5F8C) & ChrW(&H518D) & ChrW(&H91CD) & ChrW(&H65B0) & ChrW(&H57F7) & ChrW(&H884C) & _
                       ChrW(&H3002), _
                   vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
            Exit Function
        End If

        If isVisible = msoTrue Then
            presCount = app.Presentations.Count
            If Err.Number = 0 Then
                For i = 1 To presCount
                    If app.Presentations(i).Saved <> msoTrue Then
                        unsavedList = unsavedList & "    " & _
                                      app.Presentations(i).Name & vbCrLf
                    End If
                    If Err.Number <> 0 Then Exit For
                Next
            End If
            Err.Clear
            Set app = Nothing
            On Error GoTo 0

            If unsavedList <> "" Then
                MsgBox "PowerPoint " & ChrW(&H6709) & ChrW(&H5C1A) & ChrW(&H672A) & ChrW(&H5132) & ChrW(&H5B58) & ChrW(&H7684) & ChrW(&H7C21) & ChrW(&H5831) & ChrW(&HFF1A) & vbCrLf & _
                       unsavedList & vbCrLf & _
                       ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H5132) & ChrW(&H5B58) & ChrW(&H4F60) & ChrW(&H7684) & ChrW(&H5DE5) & ChrW(&H4F5C) & ChrW(&H3002) & ChrW(&H672A) & ChrW(&H505A) & ChrW(&H4EFB) & ChrW(&H4F55) & ChrW(&H66F4) & _
                           ChrW(&H6539) & ChrW(&H3002), _
                       vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
                Exit Function
            End If

            notes = notes & "- PowerPoint " & ChrW(&H5DF2) & ChrW(&H958B) & ChrW(&H555F) & ChrW(&HFF08) & ChrW(&H5168) & ChrW(&H90E8) & ChrW(&H5DF2) & ChrW(&H5132) & ChrW(&H5B58) & ChrW(&HFF09) & ChrW(&H3002) & vbCrLf
        Else
            presCount = app.Presentations.Count
            If Err.Number = 0 Then
                For i = presCount To 1 Step -1
                    app.Presentations(i).Saved = msoTrue
                    app.Presentations(i).Close
                Next
            End If
            app.Quit
            Err.Clear
            Set app = Nothing
            On Error GoTo 0
            notes = notes & "- " & ChrW(&H5DF2) & ChrW(&H79FB) & ChrW(&H9664) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H96B1) & ChrW(&H5F62) & ChrW(&H7684) & ChrW(&H5B64) & ChrW(&H7ACB) & " PowerPoint" & _
                    " " & ChrW(&H5BE6) & ChrW(&H4F8B) & ChrW(&H3002) & vbCrLf
        End If
    End If

    On Error Resume Next
    sz = fso.GetFile(lyricsPath).Size
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox ChrW(&H7121) & ChrW(&H6CD5) & ChrW(&H8B80) & ChrW(&H53D6) & " lyrics.txt" & ChrW(&H3002), vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
        Exit Function
    End If
    On Error GoTo 0

    If sz < GUARD_MIN_TEXT_BYTES Then
        MsgBox "lyrics.txt " & ChrW(&H662F) & ChrW(&H7A7A) & ChrW(&H7684) & ChrW(&HFF08) & ChrW(&H6216) & ChrW(&H53EA) & ChrW(&H6709) & " BOM" & ChrW(&HFF09) & ChrW(&H3002) & ChrW(&H8ACB) & ChrW(&H5148) & ChrW(&H586B) & ChrW(&H5165) & _
            ChrW(&H5167) & ChrW(&H5BB9) & ChrW(&H3002), _
               vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
        Exit Function
    ElseIf sz > GUARD_MAX_TEXT_BYTES Then
        MsgBox "lyrics.txt " & ChrW(&H9AD4) & ChrW(&H7A4D) & ChrW(&H7570) & ChrW(&H5E38) & ChrW(&H5927) & ChrW(&HFF08) & sz & " " & ChrW(&H4F4D) & ChrW(&H5143) & ChrW(&H7D44) & ChrW(&HFF09) & ChrW(&H3002) & _
               ChrW(&H662F) & ChrW(&H4E0D) & ChrW(&H662F) & ChrW(&H653E) & ChrW(&H932F) & ChrW(&H6A94) & ChrW(&H6848) & ChrW(&HFF1F) & ChrW(&H672A) & ChrW(&H505A) & ChrW(&H4EFB) & ChrW(&H4F55) & ChrW(&H66F4) & ChrW(&H6539) & ChrW(&H3002), _
               vbCritical, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
        Exit Function
    End If

    If notes = "" Then notes = "- PowerPoint " & ChrW(&H72C0) & ChrW(&H614B) & ChrW(&H6B63) & ChrW(&H5E38) & ChrW(&H3002) & vbCrLf

    MsgBox ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H524D) & ChrW(&H6AA2) & ChrW(&H67E5) & ChrW(&H901A) & ChrW(&H904E) & ChrW(&HFF1A) & vbCrLf & vbCrLf & notes & vbCrLf & _
           "lyrics.txt " & ChrW(&H5DF2) & ChrW(&H9A57) & ChrW(&H8B49) & ChrW(&H3002) & ChrW(&H73FE) & ChrW(&H5728) & ChrW(&H958B) & ChrW(&H59CB) & ChrW(&H3002), _
           vbInformation, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)

    PreRunGuardLyrics = True
End Function

' =====================================================================
' [ADDED] Embedded end-of-run cleanup (graceful quit + triple-gated
' orphan kill + lock-file sweep). Popup only when something was done.
' =====================================================================

Sub EndOfRunCleanup()
    Dim wmi, actions, quitDone

    actions = ""
    quitDone = False

    Set wmi = Nothing
    On Error Resume Next
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    If Err.Number <> 0 Then
        Set wmi = Nothing
        Err.Clear
    End If
    On Error GoTo 0

    If Not wmi Is Nothing Then
        If GuardOtherScriptBusy(wmi, "create-worship-folder", "refresh-master") Then
            Exit Sub    ' another worship script is mid-run; do not interfere
        End If
    End If

    GuardGracefulQuitInvisible actions, quitDone

    If GUARD_KILL_ENABLED And Not wmi Is Nothing Then
        If quitDone Then WScript.Sleep 3000
        GuardKillAgedOrphans wmi, actions
    End If

    If Not wmi Is Nothing Then
        If GuardCountPpt(wmi) = 0 Then
            GuardRemoveLockFiles scriptFolder, actions
        End If
    End If

    If actions <> "" Then
        MsgBox ChrW(&H57F7) & ChrW(&H884C) & ChrW(&H5F8C) & ChrW(&H6E05) & ChrW(&H7406) & ChrW(&HFF1A) & vbCrLf & vbCrLf & actions, _
               vbInformation, ChrW(&H5EFA) & ChrW(&H7ACB) & ChrW(&H6B4C) & ChrW(&H8A5E) & ChrW(&H7C21) & ChrW(&H5831)
    End If
End Sub

Sub GuardGracefulQuitInvisible(ByRef actions, ByRef quitDone)
    Dim app, isVisible, presCount, i

    Set app = Nothing
    On Error Resume Next
    Set app = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Or app Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    isVisible = app.Visible
    If Err.Number <> 0 Then
        Err.Clear
        Set app = Nothing
        On Error GoTo 0
        Exit Sub                        ' busy or dying instance: leave alone
    End If

    If isVisible = msoTrue Then
        Set app = Nothing
        Err.Clear
        On Error GoTo 0
        Exit Sub                        ' user instance: never touch
    End If

    presCount = app.Presentations.Count
    If Err.Number = 0 Then
        For i = presCount To 1 Step -1
            app.Presentations(i).Saved = msoTrue
            app.Presentations(i).Close
        Next
    End If
    app.Quit
    If Err.Number = 0 Then
        actions = actions & "- " & ChrW(&H5DF2) & ChrW(&H95DC) & ChrW(&H9589) & ChrW(&H4E00) & ChrW(&H500B) & ChrW(&H96B1) & ChrW(&H5F62) & ChrW(&H7684) & ChrW(&H5B64) & ChrW(&H7ACB) & " PowerPoint " & ChrW(&H5BE6) & ChrW(&H4F8B) & _
            ChrW(&H3002) & vbCrLf
        quitDone = True
    End If

    Err.Clear
    Set app = Nothing
    On Error GoTo 0
End Sub

Sub GuardKillAgedOrphans(ByVal wmi, ByRef actions)
    Dim procs, p, cmd, ageMinutes, rc

    On Error Resume Next
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId, CommandLine, CreationDate FROM Win32_Process" & _
        " WHERE Name = 'POWERPNT.EXE'")

    For Each p In procs
        cmd = "" & p.CommandLine
        If InStr(1, cmd, "embedding", vbTextCompare) > 0 _
           Or InStr(1, cmd, "/automation", vbTextCompare) > 0 Then

            ageMinutes = DateDiff("n", GuardCimToDate("" & p.CreationDate), Now)

            If Err.Number = 0 And ageMinutes >= GUARD_ORPHAN_MIN_AGE_MIN Then
                rc = p.Terminate(0)
                If rc = 0 Then
                    actions = actions & "- " & ChrW(&H5DF2) & ChrW(&H5F37) & ChrW(&H5236) & ChrW(&H7D50) & ChrW(&H675F) & ChrW(&H5B64) & ChrW(&H7ACB) & ChrW(&H7684) & " POWERPNT.EXE" & ChrW(&HFF08) & "PID " & _
                              p.ProcessId & ChrW(&HFF0C) & ChrW(&H5DF2) & ChrW(&H5B58) & ChrW(&H5728) & " " & ageMinutes & " " & ChrW(&H5206) & ChrW(&H9418) & ChrW(&HFF09) & ChrW(&H3002) & vbCrLf
                End If
            End If
            Err.Clear
        End If
    Next

    Err.Clear
    On Error GoTo 0
End Sub

Function GuardOtherScriptBusy(ByVal wmi, ByVal nameA, ByVal nameB)
    Dim procs, p, cmd

    GuardOtherScriptBusy = False

    On Error Resume Next
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId, CommandLine FROM Win32_Process" & _
        " WHERE Name = 'wscript.exe' OR Name = 'cscript.exe'")

    For Each p In procs
        cmd = "" & p.CommandLine
        If InStr(1, cmd, nameA, vbTextCompare) > 0 _
           Or InStr(1, cmd, nameB, vbTextCompare) > 0 Then
            GuardOtherScriptBusy = True
            Exit Function
        End If
    Next

    Err.Clear
    On Error GoTo 0
End Function

Function GuardCountPpt(ByVal wmi)
    Dim procs, p, n

    n = 0
    On Error Resume Next
    Set procs = wmi.ExecQuery( _
        "SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'")
    For Each p In procs
        n = n + 1
    Next
    Err.Clear
    On Error GoTo 0

    GuardCountPpt = n
End Function

Sub GuardRemoveLockFiles(ByVal rootPath, ByRef actions)
    Dim removed, childFolder

    removed = 0
    GuardRemoveLockFilesIn rootPath, removed

    On Error Resume Next
    For Each childFolder In fso.GetFolder(rootPath).SubFolders
        GuardRemoveLockFilesIn childFolder.Path, removed
    Next
    Err.Clear
    On Error GoTo 0

    If removed > 0 Then
        actions = actions & "- " & ChrW(&H5DF2) & ChrW(&H79FB) & ChrW(&H9664) & " " & removed & " " & ChrW(&H500B) & ChrW(&H6B98) & ChrW(&H7559) & ChrW(&H7684) & " ~$ " & ChrW(&H9396) & ChrW(&H5B9A) & ChrW(&H6A94) & ChrW(&H3002) & vbCrLf
    End If
End Sub

Sub GuardRemoveLockFilesIn(ByVal folderPath, ByRef removed)
    Dim f, ext, lockPaths(31), n, i

    n = 0
    On Error Resume Next
    For Each f In fso.GetFolder(folderPath).Files
        If Left(f.Name, 2) = "~$" And n <= UBound(lockPaths) Then
            ext = LCase(fso.GetExtensionName(f.Name))
            If ext = "pptx" Or ext = "ppsx" Or ext = "pptm" Then
                lockPaths(n) = f.Path
                n = n + 1
            End If
        End If
    Next
    Err.Clear

    For i = 0 To n - 1
        fso.DeleteFile lockPaths(i), True
        If Err.Number = 0 Then removed = removed + 1
        Err.Clear
    Next
    On Error GoTo 0
End Sub

Function GuardCimToDate(ByVal s)
    GuardCimToDate = DateSerial(CInt(Mid(s, 1, 4)), CInt(Mid(s, 5, 2)), CInt(Mid(s, 7, 2))) + _
                     TimeSerial(CInt(Mid(s, 9, 2)), CInt(Mid(s, 11, 2)), CInt(Mid(s, 13, 2)))
End Function