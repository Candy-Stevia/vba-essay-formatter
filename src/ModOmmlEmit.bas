Attribute VB_Name = "ModOmmlEmit"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModOmmlEmit
' Purpose: OpenXML OMML emitter and Flat OPC wrapper
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private Const MAX_EMIT_DEPTH As Long = 64

Private mParts() As String
Private mPartCount As Long
Private mEmitDepth As Long

Private Sub BufInit()
    ReDim mParts(1 To 256)
    mPartCount = 0
End Sub

Private Sub BufAppend(ByVal s As String)
    Dim newCap As Long
    If Len(s) = 0 Then
        Exit Sub
    End If
    mPartCount = mPartCount + 1
    If mPartCount > UBound(mParts) Then
        newCap = UBound(mParts) * 2
        ReDim Preserve mParts(1 To newCap)
    End If
    mParts(mPartCount) = s
End Sub

Private Function BufFlush() As String
    If mPartCount = 0 Then
        BufFlush = ""
    Else
        ReDim Preserve mParts(1 To mPartCount)
        BufFlush = Join(mParts, "")
    End If
    Erase mParts
    mPartCount = 0
End Function

Private Sub EmitTextElem(ByVal rawText As String)
    Dim esc As String
    esc = ModTexSymbols.EscapeXml(rawText)
    If Len(esc) > 0 Then
        If Left$(esc, 1) = " " Or Right$(esc, 1) = " " Or InStr(1, esc, "  ") > 0 Then
            BufAppend "<m:t xml:space=""preserve"">"
        Else
            BufAppend "<m:t>"
        End If
    Else
        BufAppend "<m:t>"
    End If
    BufAppend esc
    BufAppend "</m:t>"
End Sub

Private Sub EmitNode(ByVal idx As Long)
    Dim nType As Integer
    Dim c1 As Long, c2 As Long, c3 As Long, maxCols As Long, cIdx As Long
    Dim fType As String, subHide As String, supHide As String, colAlign As String
    Dim alignArr() As String
    Dim vItem As Variant, cellIdx As Variant
    Dim rowObj As Collection

    If idx <= 0 Or idx > ModTexAst.mNodeCount Then
        Exit Sub
    End If
    mEmitDepth = mEmitDepth + 1
    If mEmitDepth > MAX_EMIT_DEPTH Then
        ModTexAst.SetAmbiguous True
        mEmitDepth = mEmitDepth - 1
        Exit Sub
    End If

    nType = ModTexAst.mNodes(idx).NodeType
    c1 = ModTexAst.mNodes(idx).Child1
    c2 = ModTexAst.mNodes(idx).Child2
    c3 = ModTexAst.mNodes(idx).Child3

    Select Case nType
        Case AST_CHAR
            If ModTexAst.mNodes(idx).Bool1 Then
                BufAppend "<m:r><m:rPr><m:sty m:val=""p"" /></m:rPr>"
            Else
                BufAppend "<m:r>"
            End If
            EmitTextElem ModTexAst.mNodes(idx).Str1
            BufAppend "</m:r>"

        Case AST_TEXT
            If ModTexAst.mNodes(idx).Str2 = "double-struck" Then
                BufAppend "<m:r><m:rPr><m:scr m:val=""double-struck"" /><m:sty m:val=""p"" /></m:rPr>"
            ElseIf ModTexAst.mNodes(idx).Str2 = "script" Then
                BufAppend "<m:r><m:rPr><m:scr m:val=""script"" /><m:sty m:val=""p"" /></m:rPr>"
            ElseIf ModTexAst.mNodes(idx).Str2 = "bold" Then
                BufAppend "<m:r><m:rPr><m:sty m:val=""b"" /></m:rPr>"
            ElseIf ModTexAst.mNodes(idx).Str2 = "roman" Then
                BufAppend "<m:r><m:rPr><m:sty m:val=""p"" /></m:rPr>"
            Else
                BufAppend "<m:r><m:rPr><m:nor /></m:rPr>"
            End If
            EmitTextElem ModTexAst.mNodes(idx).Str1
            BufAppend "</m:r>"

        Case AST_FRAC
            If ModTexAst.mNodes(idx).Bool1 Then
                fType = "bar"
            Else
                fType = "noBar"
            End If
            BufAppend "<m:f><m:fPr><m:type m:val=""" & fType & """ /></m:fPr><m:num>"
            EmitNode c1
            BufAppend "</m:num><m:den>"
            EmitNode c2
            BufAppend "</m:den></m:f>"

        Case AST_RAD
            If c2 <> 0 Then
                BufAppend "<m:rad><m:deg>"
                EmitNode c2
                BufAppend "</m:deg><m:e>"
                EmitNode c1
                BufAppend "</m:e></m:rad>"
            Else
                BufAppend "<m:rad><m:radPr><m:degHide m:val=""on"" /></m:radPr><m:deg /><m:e>"
                EmitNode c1
                BufAppend "</m:e></m:rad>"
            End If

        Case AST_SUBSUP
            If c2 <> 0 And c3 <> 0 Then
                BufAppend "<m:sSubSup><m:e>"
                EmitNode c1
                BufAppend "</m:e><m:sub>"
                EmitNode c2
                BufAppend "</m:sub><m:sup>"
                EmitNode c3
                BufAppend "</m:sup></m:sSubSup>"
            ElseIf c2 <> 0 Then
                BufAppend "<m:sSub><m:e>"
                EmitNode c1
                BufAppend "</m:e><m:sub>"
                EmitNode c2
                BufAppend "</m:sub></m:sSub>"
            ElseIf c3 <> 0 Then
                BufAppend "<m:sSup><m:e>"
                EmitNode c1
                BufAppend "</m:e><m:sup>"
                EmitNode c3
                BufAppend "</m:sup></m:sSup>"
            Else
                EmitNode c1
            End If

        Case AST_NARY
            If c1 <> 0 Then
                subHide = "off"
            Else
                subHide = "on"
            End If
            If c2 <> 0 Then
                supHide = "off"
            Else
                supHide = "on"
            End If
            BufAppend "<m:nary><m:naryPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(ModTexAst.mNodes(idx).Str1) & """ /><m:limLoc m:val=""" & ModTexAst.mNodes(idx).Str2 & """ /><m:subHide m:val=""" & subHide & """ /><m:supHide m:val=""" & supHide & """ /></m:naryPr><m:sub>"
            If c1 <> 0 Then
                EmitNode c1
            End If
            BufAppend "</m:sub><m:sup>"
            If c2 <> 0 Then
                EmitNode c2
            End If
            BufAppend "</m:sup><m:e>"
            If c3 <> 0 Then
                EmitNode c3
            End If
            BufAppend "</m:e></m:nary>"

        Case AST_DELIM
            BufAppend "<m:d><m:dPr><m:begChr m:val=""" & ModTexSymbols.EscapeXml(ModTexAst.mNodes(idx).Str1) & """ /><m:sepChr m:val="""" /><m:endChr m:val=""" & ModTexSymbols.EscapeXml(ModTexAst.mNodes(idx).Str2) & """ />"
            If ModTexAst.mNodes(idx).Bool1 Then
                BufAppend "<m:grow />"
            End If
            BufAppend "</m:dPr><m:e>"
            EmitNode c1
            BufAppend "</m:e></m:d>"

        Case AST_MATRIX
            If ModTexAst.mNodes(idx).ChildList Is Nothing Then
                ModTexAst.SetAmbiguous True
                mEmitDepth = mEmitDepth - 1
                Exit Sub
            End If
            If ModTexAst.mNodes(idx).ChildList.Count = 0 Then
                ModTexAst.SetAmbiguous True
                mEmitDepth = mEmitDepth - 1
                Exit Sub
            End If
            alignArr = Split(ModTexAst.mNodes(idx).Str1, ",")
            maxCols = 0
            For Each vItem In ModTexAst.mNodes(idx).ChildList
                Set rowObj = vItem
                If rowObj.Count > maxCols Then
                    maxCols = rowObj.Count
                End If
            Next vItem
            If maxCols = 0 Then
                maxCols = 1
            End If
            BufAppend "<m:m><m:mPr><m:baseJc m:val=""center"" /><m:plcHide m:val=""on"" /><m:mcs>"
            For cIdx = 0 To maxCols - 1
                If cIdx <= UBound(alignArr) Then
                    colAlign = alignArr(cIdx)
                Else
                    colAlign = "center"
                End If
                BufAppend "<m:mc><m:mcPr><m:count m:val=""1"" /><m:mcJc m:val=""" & colAlign & """ /></m:mcPr></m:mc>"
            Next cIdx
            BufAppend "</m:mcs></m:mPr>"
            For Each vItem In ModTexAst.mNodes(idx).ChildList
                Set rowObj = vItem
                BufAppend "<m:mr>"
                For Each cellIdx In rowObj
                    BufAppend "<m:e>"
                    EmitNode CLng(cellIdx)
                    BufAppend "</m:e>"
                Next cellIdx
                For cIdx = rowObj.Count + 1 To maxCols
                    BufAppend "<m:e />"
                Next cIdx
                BufAppend "</m:mr>"
            Next vItem
            BufAppend "</m:m>"

        Case AST_ACC
            BufAppend "<m:acc><m:accPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(ModTexAst.mNodes(idx).Str1) & """ /></m:accPr><m:e>"
            EmitNode c1
            BufAppend "</m:e></m:acc>"

        Case AST_LIMLOW
            If ModTexAst.mNodes(idx).Str1 = "underbrace" Then
                BufAppend "<m:limLow><m:e><m:limLow><m:e>"
                EmitNode c1
                BufAppend "</m:e><m:lim><m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>" & ChrW(&H23DF) & "</m:t></m:r></m:lim></m:limLow></m:e><m:lim>"
                EmitNode c2
                BufAppend "</m:lim></m:limLow>"
            Else
                BufAppend "<m:limLow><m:e>"
                EmitNode c1
                BufAppend "</m:e><m:lim>"
                EmitNode c2
                BufAppend "</m:lim></m:limLow>"
            End If

        Case AST_LIMUPP
            BufAppend "<m:limUpp><m:e>"
            EmitNode c1
            BufAppend "</m:e><m:lim>"
            EmitNode c2
            BufAppend "</m:lim></m:limUpp>"

        Case AST_BORDERBOX
            BufAppend "<m:borderBox><m:e>"
            EmitNode c1
            BufAppend "</m:e></m:borderBox>"

        Case AST_BAR
            BufAppend "<m:bar><m:barPr><m:pos m:val=""" & ModTexAst.mNodes(idx).Str1 & """ /></m:barPr><m:e>"
            EmitNode c1
            BufAppend "</m:e></m:bar>"

        Case AST_GROUPCHR
            BufAppend "<m:groupChr><m:groupChrPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(ModTexAst.mNodes(idx).Str1) & """ /><m:pos m:val=""" & ModTexAst.mNodes(idx).Str2 & """ />"
            If ModTexAst.mNodes(idx).Str2 = "top" Then
                BufAppend "<m:vertJc m:val=""bot"" />"
            Else
                BufAppend "<m:vertJc m:val=""top"" />"
            End If
            BufAppend "</m:groupChrPr><m:e>"
            EmitNode c1
            BufAppend "</m:e></m:groupChr>"

        Case AST_LIST
            If Not ModTexAst.mNodes(idx).ChildList Is Nothing Then
                For Each vItem In ModTexAst.mNodes(idx).ChildList
                    EmitNode CLng(vItem)
                Next vItem
            End If
    End Select

    mEmitDepth = mEmitDepth - 1
End Sub

Public Function TranspileLatexToOmml(ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef outIsAmbiguous As Boolean, Optional ByRef outEqTag As String = "") As String
    Dim rootIdx As Long

    On Error GoTo ErrHandler

    rootIdx = ModTexParser.ParseLatex(rawLatex, isBlock, rootIdx, outIsAmbiguous, outEqTag)
    If outIsAmbiguous Then
        ModTexAst.ReleaseArena
        TranspileLatexToOmml = ""
        Exit Function
    End If

    mEmitDepth = 0
    BufInit
    BufAppend "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">"
    EmitNode rootIdx
    BufAppend "</m:oMath>"

    If ModTexAst.mIsAmbiguous Then
        outIsAmbiguous = True
        Erase mParts
        mPartCount = 0
        ModTexAst.ReleaseArena
        TranspileLatexToOmml = ""
        Exit Function
    End If

    TranspileLatexToOmml = BufFlush()
    ModTexAst.ReleaseArena
    Exit Function

ErrHandler:
    outIsAmbiguous = True
    Erase mParts
    mPartCount = 0
    ModTexAst.ReleaseArena
    TranspileLatexToOmml = ""
End Function

Public Function WrapOmmlInFlatOpc(ByVal ommlXml As String) As String
    Dim lines(0 To 21) As String
    lines(0) = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
    lines(1) = "<?mso-application progid=""Word.Document""?>"
    lines(2) = "<pkg:package xmlns:pkg=""http://schemas.microsoft.com/office/2006/xmlPackage"">"
    lines(3) = "  <pkg:part pkg:name=""/_rels/.rels"" pkg:contentType=""application/vnd.openxmlformats-package.relationships+xml"">"
    lines(4) = "    <pkg:xmlData>"
    lines(5) = "      <Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">"
    lines(6) = "        <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>"
    lines(7) = "      </Relationships>"
    lines(8) = "    </pkg:xmlData>"
    lines(9) = "  </pkg:part>"
    lines(10) = "  <pkg:part pkg:name=""/word/document.xml"" pkg:contentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"">"
    lines(11) = "    <pkg:xmlData>"
    lines(12) = "      <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">"
    lines(13) = "        <w:body>"
    lines(14) = "          <w:p>"
    lines(15) = "            " & ommlXml
    lines(16) = "          </w:p>"
    lines(17) = "        </w:body>"
    lines(18) = "      </w:document>"
    lines(19) = "    </pkg:xmlData>"
    lines(20) = "  </pkg:part>"
    lines(21) = "</pkg:package>"
    WrapOmmlInFlatOpc = Join(lines, vbCrLf)
End Function

Public Function TranspileLatexForTest(ByVal rawLatex As String, ByVal isBlock As Boolean) As String
    Dim isAmb As Boolean
    Dim eqTag As String
    Dim xml As String
    xml = TranspileLatexToOmml(rawLatex, isBlock, isAmb, eqTag)
    If isAmb Then
        TranspileLatexForTest = "AMBIGUOUS_PINK||"
    Else
        TranspileLatexForTest = "OK|" & eqTag & "|" & xml
    End If
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
