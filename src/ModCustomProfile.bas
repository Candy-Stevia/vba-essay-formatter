Attribute VB_Name = "ModCustomProfile"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ModCustomProfile: AF-PATCH v2 @STYLE custom profile & @TAG (bm_af_*) cache (< 400 lines, 7-bit ASCII)
Public Type CustomElemStyle
    HasFont As Boolean
    FontName As String
    HasSize As Boolean
    FontSize As Single
    HasColor As Boolean
    FontColor As Long
    HasBold As Boolean
    IsBold As Boolean
    HasItalic As Boolean
    IsItalic As Boolean
    HasUnderline As Boolean
    IsUnderline As Boolean
    HasUpper As Boolean
    IsUpper As Boolean
    HasKeepWithNext As Boolean
    KeepWithNext As Boolean
    HasAlign As Boolean
    Alignment As Long
    HasFirstIndent As Boolean
    FirstIndentPts As Single
    HasLeftIndent As Boolean
    LeftIndentPts As Single
    HasSpaceBefore As Boolean
    SpaceBeforePts As Single
    HasSpaceAfter As Boolean
    SpaceAfterPts As Single
End Type

Public Const TAG_CODE_SKIP As Long = -1
Public g_ForceVisualReset As Boolean
Private Const STYLE_VAR_PREFIX As String = "af_style_"

Private mStyles(1 To 8) As CustomElemStyle
Private mHasStyle(1 To 8) As Boolean
Private mHasAnyStyle As Boolean
Private mActiveSummary As String

Private Function ScopeKeyToIndex(ByVal scopeKey As String) As Long
    Dim keys As Variant, u As String, i As Long
    u = UCase$(Trim$(scopeKey))
    keys = Array("L1", "L2", "L3", "L4", "TBL", "FIG", "SRC", "BODY")
    For i = 0 To 7
        If CStr(keys(i)) = u Then
            ScopeKeyToIndex = i + 1
            Exit Function
        End If
    Next i
End Function

Private Function IndexToScopeKey(ByVal idx As Long) As String
    Dim keys As Variant
    keys = Array("", "L1", "L2", "L3", "L4", "TBL", "FIG", "SRC", "BODY")
    If idx >= 1 And idx <= 8 Then IndexToScopeKey = CStr(keys(idx))
End Function

Public Function IsValidTagType(ByVal tagType As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(tagType))
    If u = "SKIP" Or u = "CLEAR" Then
        IsValidTagType = True
    Else
        IsValidTagType = (ScopeKeyToIndex(u) >= 1 And ScopeKeyToIndex(u) <= 7)
    End If
End Function

Public Function TagTypeToCode(ByVal tagType As String) As Long
    Dim u As String
    u = UCase$(Trim$(tagType))
    If u = "SKIP" Then TagTypeToCode = TAG_CODE_SKIP Else TagTypeToCode = ScopeKeyToIndex(u)
End Function

Private Function ParseStrictNumber(ByVal s As String, ByVal minVal As Single, ByVal maxVal As Single, ByRef outVal As Single) As Boolean
    Dim t As String, k As Long, ch As Long, dotCnt As Long, hasDigit As Boolean
    t = Trim$(s)
    If Len(t) = 0 Or Len(t) > 12 Then Exit Function
    For k = 1 To Len(t)
        ch = AscW(Mid$(t, k, 1))
        If ch = 46 Then
            dotCnt = dotCnt + 1
            If dotCnt > 1 Then Exit Function
        ElseIf ch >= 48 And ch <= 57 Then
            hasDigit = True
        Else
            Exit Function
        End If
    Next k
    If Not hasDigit Then Exit Function
    Dim v As Single
    v = CSng(Val(t))
    If v < minVal Or v > maxVal Then Exit Function
    outVal = v
    ParseStrictNumber = True
End Function

Private Function ParseHexByte(ByVal h2 As String, ByRef outByte As Long) As Boolean
    If Len(h2) <> 2 Then Exit Function
    Dim k As Long, ch As Long
    For k = 1 To 2
        ch = AscW(Mid$(h2, k, 1))
        If Not ((ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 70) Or (ch >= 97 And ch <= 102)) Then Exit Function
    Next k
    outByte = CLng("&H" & h2)
    ParseHexByte = True
End Function

Private Function ParseColorSpec(ByVal s As String, ByRef outColor As Long, ByRef outNorm As String) As Boolean
    Dim t As String, r As Long, g As Long, b As Long
    t = LCase$(Trim$(s))
    If t = "auto" Then
        outColor = -16777216
        outNorm = "auto"
        ParseColorSpec = True
        Exit Function
    End If
    If Len(t) <> 7 Or Left$(t, 1) <> "#" Then Exit Function
    If Not ParseHexByte(Mid$(t, 2, 2), r) Then Exit Function
    If Not ParseHexByte(Mid$(t, 4, 2), g) Then Exit Function
    If Not ParseHexByte(Mid$(t, 6, 2), b) Then Exit Function
    outColor = RGB(r, g, b)
    outNorm = UCase$(t)
    ParseColorSpec = True
End Function

Private Function ParseBoolProp(ByVal v As String, ByVal keyName As String, ByRef hasFlag As Boolean, _
                               ByRef outFlag As Boolean, ByRef outPair As String, ByRef outErr As String) As Boolean
    Dim t As String
    t = Trim$(v)
    If t <> "0" And t <> "1" Then
        outErr = keyName & " must be 0 or 1"
        Exit Function
    End If
    hasFlag = True
    outFlag = (t = "1")
    outPair = keyName & "=" & t
    ParseBoolProp = True
End Function

Private Function ParseNumProp(ByVal v As String, ByVal keyName As String, ByVal minV As Single, ByVal maxV As Single, _
                              ByVal toCmPts As Boolean, ByRef hasFlag As Boolean, ByRef outVal As Single, _
                              ByRef outPair As String, ByRef outErr As String) As Boolean
    Dim numV As Single
    If Not ParseStrictNumber(v, minV, maxV, numV) Then
        outErr = keyName & " out of range (" & minV & "-" & maxV & ")"
        Exit Function
    End If
    hasFlag = True
    If toCmPts Then outVal = ModConstants.CmToPoints(numV) Else outVal = numV
    outPair = keyName & "=" & Trim$(Str$(numV))
    ParseNumProp = True
End Function

Private Function IsValidFontName(ByVal s As String) As Boolean
    Dim t As String, k As Long, ch As Long
    t = Trim$(s)
    If Len(t) < 1 Or Len(t) > 40 Then Exit Function
    For k = 1 To Len(t)
        ch = AscW(Mid$(t, k, 1))
        If ch < 32 Or ch = 124 Or ch = 61 Or ch = 59 Or ch = 60 Or ch = 62 Then Exit Function
    Next k
    IsValidFontName = True
End Function

Private Function ParseStyleSpecInto(ByVal specStr As String, ByRef outSt As CustomElemStyle, ByRef outNorm As String, ByRef outErr As String) As Boolean
    Dim emptySt As CustomElemStyle
    outSt = emptySt
    outNorm = ""
    Dim parts() As String, i As Long, eqPos As Long, k As String, v As String, colNorm As String, normPair As String, ok As Boolean
    parts = Split(specStr, "|")
    For i = LBound(parts) To UBound(parts)
        Dim tok As String
        tok = Trim$(parts(i))
        If Len(tok) = 0 Then GoTo NextTok
        eqPos = InStr(1, tok, "=", vbBinaryCompare)
        If eqPos <= 1 Or eqPos >= Len(tok) Then
            outErr = "Invalid key=value token: " & tok
            Exit Function
        End If
        k = LCase$(Trim$(Left$(tok, eqPos - 1)))
        v = Trim$(Mid$(tok, eqPos + 1))
        normPair = ""
        ok = False
        Select Case k
            Case "font"
                If IsValidFontName(v) Then
                    outSt.HasFont = True
                    outSt.FontName = v
                    normPair = "font=" & v
                    ok = True
                Else
                    outErr = "Invalid font name"
                End If
            Case "size"
                ok = ParseNumProp(v, "size", 8#, 36#, False, outSt.HasSize, outSt.FontSize, normPair, outErr)
            Case "color"
                If ParseColorSpec(v, outSt.FontColor, colNorm) Then
                    outSt.HasColor = True
                    normPair = "color=" & colNorm
                    ok = True
                Else
                    outErr = "Invalid color (#RRGGBB or auto)"
                End If
            Case "bold"
                ok = ParseBoolProp(v, "bold", outSt.HasBold, outSt.IsBold, normPair, outErr)
            Case "italic"
                ok = ParseBoolProp(v, "italic", outSt.HasItalic, outSt.IsItalic, normPair, outErr)
            Case "underline"
                ok = ParseBoolProp(v, "underline", outSt.HasUnderline, outSt.IsUnderline, normPair, outErr)
            Case "upper"
                ok = ParseBoolProp(v, "upper", outSt.HasUpper, outSt.IsUpper, normPair, outErr)
            Case "keepwithnext"
                ok = ParseBoolProp(v, "keepWithNext", outSt.HasKeepWithNext, outSt.KeepWithNext, normPair, outErr)
            Case "align"
                Dim av As String
                av = LCase$(v)
                If av = "left" Or av = "center" Or av = "right" Or av = "justify" Then
                    If av = "left" Then outSt.Alignment = 0
                    If av = "center" Then outSt.Alignment = 1
                    If av = "right" Then outSt.Alignment = 2
                    If av = "justify" Then outSt.Alignment = 3
                    outSt.HasAlign = True
                    normPair = "align=" & av
                    ok = True
                Else
                    outErr = "align must be left/center/right/justify"
                End If
            Case "firstindent"
                ok = ParseNumProp(v, "firstIndent", 0#, 5#, True, outSt.HasFirstIndent, outSt.FirstIndentPts, normPair, outErr)
            Case "leftindent"
                ok = ParseNumProp(v, "leftIndent", 0#, 5#, True, outSt.HasLeftIndent, outSt.LeftIndentPts, normPair, outErr)
            Case "spacebefore"
                ok = ParseNumProp(v, "spaceBefore", 0#, 72#, False, outSt.HasSpaceBefore, outSt.SpaceBeforePts, normPair, outErr)
            Case "spaceafter"
                ok = ParseNumProp(v, "spaceAfter", 0#, 72#, False, outSt.HasSpaceAfter, outSt.SpaceAfterPts, normPair, outErr)
            Case Else
                outErr = "Unknown @STYLE key: " & k
        End Select
        If Not ok Then Exit Function
        If Len(outNorm) > 0 Then outNorm = outNorm & "|"
        outNorm = outNorm & normPair
NextTok:
    Next i
    If Len(outNorm) = 0 Then
        outErr = "No valid @STYLE properties"
        Exit Function
    End If
    ParseStyleSpecInto = True
End Function

Public Function ValidateStyleDirective(ByVal rawLine As String, ByRef outScope As String, _
                                       ByRef outNormSpec As String, ByRef outIsReset As Boolean, _
                                       ByRef outErr As String) As Boolean
    outScope = ""
    outNormSpec = ""
    outIsReset = False
    outErr = ""
    Dim pipePos As Long, scopeRaw As String, restSpec As String
    pipePos = InStr(1, rawLine, "|", vbBinaryCompare)
    If pipePos <= 1 Or pipePos >= Len(rawLine) Then
        outErr = "Missing '|' separator"
        Exit Function
    End If
    scopeRaw = UCase$(Trim$(Left$(rawLine, pipePos - 1)))
    restSpec = Trim$(Mid$(rawLine, pipePos + 1))
    If scopeRaw = "RESET" Then
        Dim rstArg As String
        rstArg = UCase$(restSpec)
        If rstArg = "ALL" Or ScopeKeyToIndex(rstArg) > 0 Then
            outIsReset = True
            outScope = rstArg
            ValidateStyleDirective = True
        Else
            outErr = "Invalid RESET target: " & restSpec
        End If
        Exit Function
    End If
    If ScopeKeyToIndex(scopeRaw) = 0 Then
        outErr = "Unknown @STYLE scope: " & scopeRaw
        Exit Function
    End If
    Dim parsedSt As CustomElemStyle
    If Not ParseStyleSpecInto(restSpec, parsedSt, outNormSpec, outErr) Then Exit Function
    outScope = scopeRaw
    ValidateStyleDirective = True
End Function

Private Function MergeStyleSpecs(ByVal oldSpec As String, ByVal newSpec As String) As String
    If Len(Trim$(oldSpec)) = 0 Then
        MergeStyleSpecs = newSpec
        Exit Function
    End If
    Dim dict As Object, orderKeys() As String, kCnt As Long, i As Long
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Dim specs(1 To 2) As String, sIdx As Long, parts() As String, tok As String, eqPos As Long, k As String
    specs(1) = oldSpec
    specs(2) = newSpec
    For sIdx = 1 To 2
        parts = Split(specs(sIdx), "|")
        For i = LBound(parts) To UBound(parts)
            tok = Trim$(parts(i))
            eqPos = InStr(1, tok, "=", vbBinaryCompare)
            If eqPos > 1 And eqPos < Len(tok) Then
                k = LCase$(Trim$(Left$(tok, eqPos - 1)))
                If Not dict.Exists(k) Then
                    kCnt = kCnt + 1
                    ReDim Preserve orderKeys(1 To kCnt)
                    orderKeys(kCnt) = k
                End If
                dict(k) = tok
            End If
        Next i
    Next sIdx
    Dim res As String
    For i = 1 To kCnt
        If Len(res) > 0 Then res = res & "|"
        res = res & CStr(dict(orderKeys(i)))
    Next i
    MergeStyleSpecs = res
End Function

Public Sub ApplyStyleDirective(ByVal doc As Document, ByVal scopeKey As String, ByVal normSpec As String, ByVal isReset As Boolean)
    On Error Resume Next
    If isReset Then
        If UCase$(scopeKey) = "ALL" Then
            Dim v As Variable, names() As String, cnt As Long, k As Long
            For Each v In doc.Variables
                If Left$(v.Name, Len(STYLE_VAR_PREFIX)) = STYLE_VAR_PREFIX Then
                    cnt = cnt + 1
                    ReDim Preserve names(1 To cnt)
                    names(cnt) = v.Name
                End If
            Next v
            For k = 1 To cnt
                doc.Variables(names(k)).Delete
            Next k
        Else
            doc.Variables(STYLE_VAR_PREFIX & UCase$(scopeKey)).Delete
        End If
        Err.Clear
    Else
        Dim varName As String, prevSpec As String, mergedSpec As String
        varName = STYLE_VAR_PREFIX & UCase$(scopeKey)
        prevSpec = ""
        prevSpec = doc.Variables(varName).Value
        Err.Clear
        mergedSpec = MergeStyleSpecs(prevSpec, normSpec)
        ModMathHelpers.SetDocVariable doc, varName, mergedSpec
    End If
End Sub

Private Sub ClearStyleArrays()
    Dim i As Long, emptySt As CustomElemStyle
    For i = 1 To 8
        mStyles(i) = emptySt
        mHasStyle(i) = False
    Next i
    mHasAnyStyle = False
    mActiveSummary = ""
End Sub

Public Sub ResetStyleProfileCache()
    ClearStyleArrays
    g_ForceVisualReset = False
End Sub

Public Sub LoadStyleProfile(ByVal doc As Document)
    ClearStyleArrays
    If doc Is Nothing Then Exit Sub
    On Error Resume Next
    Dim i As Long, sk As String, rawSpec As String, normOut As String, errOut As String
    For i = 1 To 8
        sk = IndexToScopeKey(i)
        rawSpec = ""
        rawSpec = doc.Variables(STYLE_VAR_PREFIX & sk).Value
        If Err.Number <> 0 Then
            Err.Clear
            rawSpec = ""
        End If
        If Len(rawSpec) > 0 Then
            If ParseStyleSpecInto(rawSpec, mStyles(i), normOut, errOut) Then
                mHasStyle(i) = True
                mHasAnyStyle = True
                If Len(mActiveSummary) > 0 Then mActiveSummary = mActiveSummary & ", "
                mActiveSummary = mActiveSummary & sk
            End If
        End If
    Next i
End Sub

Public Function ResolveEffectiveFont(ByVal elemType As Long, ByVal fallbackFont As String) As String
    If mHasAnyStyle And elemType >= 1 And elemType <= 8 Then
        If mHasStyle(elemType) And mStyles(elemType).HasFont Then
            ResolveEffectiveFont = mStyles(elemType).FontName
            Exit Function
        End If
    End If
    ResolveEffectiveFont = fallbackFont
End Function

Public Function HasCustomStyle(ByVal elemType As Long) As Boolean
    If mHasAnyStyle And elemType >= 1 And elemType <= 8 Then HasCustomStyle = mHasStyle(elemType)
End Function
Public Function IsL1UpperDisabled() As Boolean
    If mHasStyle(1) Then If mStyles(1).HasUpper And Not mStyles(1).IsUpper Then IsL1UpperDisabled = True
End Function
Public Function GetActiveStyleSummary() As String
    GetActiveStyleSummary = mActiveSummary
End Function

Public Sub ApplyCustomStyleOverrides(ByVal doc As Document, ByVal rng As Range, ByVal elemType As Long, _
                                     ByVal hasOMaths As Boolean, Optional ByVal skipFont As Boolean = False)
    If Not mHasAnyStyle Or elemType < 1 Or elemType > 8 Then Exit Sub
    If Not mHasStyle(elemType) Then Exit Sub
    Dim st As CustomElemStyle
    st = mStyles(elemType)
    If st.HasFont And Not skipFont Then ModFormatSingle.ApplyFontNameSkippingMath doc, rng, st.FontName, hasOMaths
    If st.HasSize Then rng.Font.Size = st.FontSize
    If st.HasColor Then rng.Font.Color = st.FontColor
    If st.HasBold Then
        rng.Font.Bold = st.IsBold
        rng.Bold = st.IsBold
    End If
    If st.HasItalic Then rng.Font.Italic = st.IsItalic
    If st.HasUnderline Then
        If st.IsUnderline Then rng.Font.Underline = 1 Else rng.Font.Underline = 0
    End If
    If st.HasUpper Then ModFormatSingle.ApplyAllCapsSkippingMath doc, rng, st.IsUpper, hasOMaths
    If st.HasAlign Then rng.ParagraphFormat.Alignment = st.Alignment
    If st.HasFirstIndent Then rng.ParagraphFormat.FirstLineIndent = st.FirstIndentPts
    If st.HasLeftIndent Then rng.ParagraphFormat.LeftIndent = st.LeftIndentPts
    If st.HasSpaceBefore Then rng.ParagraphFormat.SpaceBefore = st.SpaceBeforePts
    If st.HasSpaceAfter Then rng.ParagraphFormat.SpaceAfter = st.SpaceAfterPts
    If st.HasKeepWithNext Then rng.ParagraphFormat.KeepWithNext = st.KeepWithNext
End Sub

Private Sub CountAndClearProfileItems(ByVal doc As Document, ByVal doDeleteStyles As Boolean, _
                                      ByVal doDeleteTags As Boolean, ByRef outStyleCnt As Long, ByRef outTagCnt As Long)
    Dim v As Variable, bm As Bookmark, sNames() As String, tNames() As String, k As Long
    outStyleCnt = 0
    outTagCnt = 0
    On Error Resume Next
    For Each v In doc.Variables
        If Left$(LCase$(v.Name), Len(STYLE_VAR_PREFIX)) = STYLE_VAR_PREFIX Then
            outStyleCnt = outStyleCnt + 1
            ReDim Preserve sNames(1 To outStyleCnt)
            sNames(outStyleCnt) = v.Name
        End If
    Next v
    For Each bm In doc.Bookmarks
        If Left$(LCase$(bm.Name), 6) = "bm_af_" Then
            outTagCnt = outTagCnt + 1
            ReDim Preserve tNames(1 To outTagCnt)
            tNames(outTagCnt) = bm.Name
        End If
    Next bm
    If doDeleteStyles Then
        For k = 1 To outStyleCnt
            doc.Variables(sNames(k)).Delete
        Next k
    End If
    If doDeleteTags Then
        For k = 1 To outTagCnt
            ModZoneManager.SafeDelBm doc, tNames(k)
        Next k
    End If
    Err.Clear
End Sub

Public Function ResetCustomProfileForTest(ByVal doc As Document, ByVal clearTags As Boolean, ByVal reformatClean As Boolean) As String
    If doc Is Nothing Then
        ResetCustomProfileForTest = "ERR|NODOC"
        Exit Function
    End If
    Dim sCnt As Long, tCnt As Long
    CountAndClearProfileItems doc, True, clearTags, sCnt, tCnt
    ResetStyleProfileCache
    ResetTagCache
    If reformatClean Then
        On Error Resume Next
        If Len(Trim$(doc.Path)) > 0 And Not doc.ReadOnly And Not ModBackup.IsCloudPath(doc.FullName) Then doc.Save
        Err.Clear
        On Error GoTo 0
        g_ForceVisualReset = True
        ModRibbonCallbacks.EnsureConfigInitialized
        ModMain.RunEngine doc, EngineModeFull
        g_ForceVisualReset = False
    End If
    ResetCustomProfileForTest = "OK|" & CStr(sCnt) & "|" & CStr(IIf(clearTags, tCnt, 0))
End Function

Public Sub ResetCustomProfileInteractive(ByVal doc As Document)
    If ModMain.g_IsEngineRunning Or ModMathController.g_IsMathBuilding Then Exit Sub
    If doc Is Nothing Then Exit Sub
    Dim sCnt As Long, tCnt As Long
    CountAndClearProfileItems doc, False, False, sCnt, tCnt
    If sCnt = 0 And tCnt = 0 Then
        ModUI.MsgBoxW ModConstants.Uni("No custom @STYLE or @TAG found. Document is already using default academic style. (T\u00e0i li\u1ec7u \u0111ang d\u00f9ng chu\u1ea9n m\u1eb7c \u0111\u1ecbnh, kh\u00f4ng c\u00f3 @STYLE hay @TAG n\u00e0o.)"), _
                      vbInformation, ModConstants.Uni("Reset Custom Style (V\u1ec1 m\u1eb7c \u0111\u1ecbnh)")
        Exit Sub
    End If
    Dim qMsg As String, ans As Long, clearTags As Boolean
    qMsg = ModConstants.Uni("Active custom profile detected (Ph\u00e1t hi\u1ec7n h\u1ed3 s\u01a1 t\u00f9y ch\u1ec9nh):") & vbCrLf & _
           "- @STYLE scopes: " & CStr(sCnt) & vbCrLf & _
           "- @TAG bookmarks: " & CStr(tCnt) & vbCrLf & vbCrLf & _
           ModConstants.Uni("[Yes] Reset @STYLE only, keep @TAG (Ch\u1ec9 x\u00f3a @STYLE, gi\u1eef @TAG)") & vbCrLf & _
           ModConstants.Uni("[No] Reset BOTH @STYLE and @TAG (X\u00f3a s\u1ea1ch c\u1ea3 @STYLE v\u00e0 @TAG)") & vbCrLf & _
           ModConstants.Uni("[Cancel] Cancel (H\u1ee7y)")
    ans = ModUI.MsgBoxW(qMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1, ModConstants.Uni("Reset Custom Style (V\u1ec1 m\u1eb7c \u0111\u1ecbnh)"))
    If ans = 2 Then Exit Sub
    clearTags = (ans = 7)
    CountAndClearProfileItems doc, True, clearTags, sCnt, tCnt
    ResetStyleProfileCache
    ResetTagCache
    Dim reMsg As String, reAns As Long
    reMsg = ModConstants.Uni("Custom profile cleared. Run Format All now to restore default black/clean typography? (\u0110\u00e3 x\u00f3a h\u1ed3 s\u01a1 t\u00f9y ch\u1ec9nh. Ch\u1ea1y l\u1ea1i Format All ngay \u0111\u1ec3 t\u1ea9y m\u00e0u/font/g\u1ea1ch ch\u00e2n v\u1ec1 chu\u1ea9n m\u1eb7c \u0111\u1ecbnh?)")
    reAns = ModUI.MsgBoxW(reMsg, vbYesNo + vbQuestion + vbDefaultButton1, ModConstants.Uni("Re-Apply Default Style (\u00c1p l\u1ea1i chu\u1ea9n m\u1eb7c \u0111\u1ecbnh)"))
    If reAns = 6 Then
        On Error Resume Next
        If Len(Trim$(doc.Path)) > 0 And Not doc.ReadOnly And Not ModBackup.IsCloudPath(doc.FullName) Then doc.Save
        Err.Clear
        On Error GoTo 0
        g_ForceVisualReset = True
        ModRibbonCallbacks.EnsureConfigInitialized
        ModMain.RunEngine doc, EngineModeFull
        g_ForceVisualReset = False
    End If
End Sub

Public Sub ApplyTagToPara(ByVal doc As Document, ByVal pStart As Long, ByVal pEnd As Long, ByVal tagType As String)
    ModBookmark.ApplyAfTagBookmark doc, pStart, pEnd, tagType
End Sub
Public Sub ResetTagCache()
    ModBookmark.ResetAfTagCache
End Sub
Public Sub LoadTagCache(ByVal doc As Document)
    ModBookmark.LoadAfTagCache doc
End Sub
Public Function HasActiveTags() As Boolean
    HasActiveTags = ModBookmark.HasActiveAfTags()
End Function
Public Function LookupTagForRange(ByVal rStart As Long, ByVal rEnd As Long) As Long
    LookupTagForRange = ModBookmark.LookupAfTagForRange(rStart, rEnd)
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
