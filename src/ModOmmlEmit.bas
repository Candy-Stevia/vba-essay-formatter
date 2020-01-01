Attribute VB_Name = "ModOmmlEmit"
Option Explicit

' ==============================================================================
' Module: ModOmmlEmit
' Purpose: OpenXML OMML emitter and Flat OPC wrapper
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private Function EmitNode(ByVal idx As Long) As String
    Dim nd As AstNode
    Dim res As String
    Dim vItem As Variant
    Dim subXml As String
    Dim supXml As String
    Dim bodyXml As String
    Dim baseXml As String
    Dim limXml As String
    Dim numXml As String
    Dim denXml As String
    Dim degXml As String
    Dim fType As String
    Dim subHide As String
    Dim supHide As String
    Dim growXml As String
    Dim alignArr() As String
    Dim maxCols As Long
    Dim mcsXml As String
    Dim mprXml As String
    Dim mrXml As String
    Dim rowObj As Collection
    Dim cellIdx As Variant
    Dim cIdx As Long
    Dim colAlign As String
    Dim eXml As String

    If idx <= 0 Or idx > ModTexAst.mNodeCount Then
        EmitNode = ""
        Exit Function
    End If

    nd = ModTexAst.mNodes(idx)

    Select Case nd.NodeType
        Case AST_CHAR
            If nd.Bool1 Then
                EmitNode = "<m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            Else
                EmitNode = "<m:r><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            End If

        Case AST_TEXT
            If nd.Str2 = "double-struck" Then
                EmitNode = "<m:r><m:rPr><m:sty m:val=""p"" /><m:scr m:val=""double-struck"" /></m:rPr><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            ElseIf nd.Str2 = "script" Then
                EmitNode = "<m:r><m:rPr><m:sty m:val=""p"" /><m:scr m:val=""script"" /></m:rPr><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            ElseIf nd.Str2 = "bold" Then
                EmitNode = "<m:r><m:rPr><m:sty m:val=""b"" /></m:rPr><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            ElseIf nd.Str2 = "roman" Then
                EmitNode = "<m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            Else
                EmitNode = "<m:r><m:rPr><m:nor /><m:sty m:val=""p"" /></m:rPr><m:t>" & ModTexSymbols.EscapeXml(nd.Str1) & "</m:t></m:r>"
            End If

        Case AST_FRAC
            If nd.Bool1 Then
                fType = "bar"
            Else
                fType = "noBar"
            End If
            numXml = EmitNode(nd.Child1)
            denXml = EmitNode(nd.Child2)
            EmitNode = "<m:f><m:fPr><m:type m:val=""" & fType & """ /></m:fPr><m:num>" & numXml & "</m:num><m:den>" & denXml & "</m:den></m:f>"

        Case AST_RAD
            baseXml = EmitNode(nd.Child1)
            If nd.Child2 <> 0 Then
                degXml = EmitNode(nd.Child2)
                EmitNode = "<m:rad><m:deg>" & degXml & "</m:deg><m:e>" & baseXml & "</m:e></m:rad>"
            Else
                EmitNode = "<m:rad><m:radPr><m:degHide m:val=""on"" /></m:radPr><m:deg /><m:e>" & baseXml & "</m:e></m:rad>"
            End If

        Case AST_SUBSUP
            baseXml = EmitNode(nd.Child1)
            subXml = ""
            supXml = ""
            If nd.Child2 <> 0 Then
                subXml = EmitNode(nd.Child2)
            End If
            If nd.Child3 <> 0 Then
                supXml = EmitNode(nd.Child3)
            End If
            If nd.Child2 <> 0 And nd.Child3 <> 0 Then
                EmitNode = "<m:sSubSup><m:e>" & baseXml & "</m:e><m:sub>" & subXml & "</m:sub><m:sup>" & supXml & "</m:sup></m:sSubSup>"
            ElseIf nd.Child2 <> 0 Then
                EmitNode = "<m:sSub><m:e>" & baseXml & "</m:e><m:sub>" & subXml & "</m:sub></m:sSub>"
            ElseIf nd.Child3 <> 0 Then
                EmitNode = "<m:sSup><m:e>" & baseXml & "</m:e><m:sup>" & supXml & "</m:sup></m:sSup>"
            Else
                EmitNode = baseXml
            End If

        Case AST_NARY
            subXml = ""
            supXml = ""
            bodyXml = ""
            If nd.Child1 <> 0 Then
                subXml = EmitNode(nd.Child1)
                subHide = "off"
            Else
                subHide = "on"
            End If
            If nd.Child2 <> 0 Then
                supXml = EmitNode(nd.Child2)
                supHide = "off"
            Else
                supHide = "on"
            End If
            If nd.Child3 <> 0 Then
                bodyXml = EmitNode(nd.Child3)
            End If
            EmitNode = "<m:nary><m:naryPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(nd.Str1) & """ /><m:limLoc m:val=""" & nd.Str2 & """ /><m:subHide m:val=""" & subHide & """ /><m:supHide m:val=""" & supHide & """ /></m:naryPr><m:sub>" & subXml & "</m:sub><m:sup>" & supXml & "</m:sup><m:e>" & bodyXml & "</m:e></m:nary>"

        Case AST_DELIM
            bodyXml = EmitNode(nd.Child1)
            If nd.Bool1 Then
                growXml = "<m:grow />"
            Else
                growXml = ""
            End If
            EmitNode = "<m:d><m:dPr><m:begChr m:val=""" & ModTexSymbols.EscapeXml(nd.Str1) & """ /><m:endChr m:val=""" & ModTexSymbols.EscapeXml(nd.Str2) & """ /><m:sepChr m:val="""" />" & growXml & "</m:dPr><m:e>" & bodyXml & "</m:e></m:d>"

        Case AST_MATRIX
            alignArr = Split(nd.Str1, ",")
            maxCols = 0
            For Each vItem In nd.ChildList
                Set rowObj = vItem
                If rowObj.Count > maxCols Then
                    maxCols = rowObj.Count
                End If
            Next vItem
            If maxCols = 0 Then
                maxCols = 1
            End If
            mcsXml = ""
            For cIdx = 0 To maxCols - 1
                If cIdx <= UBound(alignArr) Then
                    colAlign = alignArr(cIdx)
                Else
                    colAlign = "center"
                End If
                mcsXml = mcsXml & "<m:mc><m:mcPr><m:mcJc m:val=""" & colAlign & """ /><m:count m:val=""1"" /></m:mcPr></m:mc>"
            Next cIdx
            mprXml = "<m:mPr><m:baseJc m:val=""center"" /><m:plcHide m:val=""on"" /><m:mcs>" & mcsXml & "</m:mcs></m:mPr>"
            mrXml = ""
            For Each vItem In nd.ChildList
                Set rowObj = vItem
                eXml = ""
                For Each cellIdx In rowObj
                    eXml = eXml & "<m:e>" & EmitNode(CLng(cellIdx)) & "</m:e>"
                Next cellIdx
                mrXml = mrXml & "<m:mr>" & eXml & "</m:mr>"
            Next vItem
            EmitNode = "<m:m>" & mprXml & mrXml & "</m:m>"

        Case AST_ACC
            baseXml = EmitNode(nd.Child1)
            EmitNode = "<m:acc><m:accPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(nd.Str1) & """ /></m:accPr><m:e>" & baseXml & "</m:e></m:acc>"

        Case AST_LIMLOW
            baseXml = EmitNode(nd.Child1)
            limXml = EmitNode(nd.Child2)
            If nd.Str1 = "underbrace" Then
                EmitNode = "<m:limLow><m:e><m:limLow><m:e>" & baseXml & "</m:e><m:lim><m:r><m:rPr><m:sty m:val=""p"" /></m:rPr><m:t>" & ChrW(&H23DF) & "</m:t></m:r></m:lim></m:limLow></m:e><m:lim>" & limXml & "</m:lim></m:limLow>"
            Else
                EmitNode = "<m:limLow><m:e>" & baseXml & "</m:e><m:lim>" & limXml & "</m:lim></m:limLow>"
            End If

        Case AST_LIMUPP
            baseXml = EmitNode(nd.Child1)
            limXml = EmitNode(nd.Child2)
            EmitNode = "<m:limUpp><m:e>" & baseXml & "</m:e><m:lim>" & limXml & "</m:lim></m:limUpp>"

        Case AST_BORDERBOX
            bodyXml = EmitNode(nd.Child1)
            EmitNode = "<m:borderBox><m:e>" & bodyXml & "</m:e></m:borderBox>"

        Case AST_BAR
            baseXml = EmitNode(nd.Child1)
            EmitNode = "<m:bar><m:barPr><m:pos m:val=""" & nd.Str1 & """ /></m:barPr><m:e>" & baseXml & "</m:e></m:bar>"

        Case AST_GROUPCHR
            baseXml = EmitNode(nd.Child1)
            If nd.Str2 = "top" Then
                EmitNode = "<m:groupChr><m:groupChrPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(nd.Str1) & """ /><m:pos m:val=""top"" /><m:vertJc m:val=""bot"" /></m:groupChrPr><m:e>" & baseXml & "</m:e></m:groupChr>"
            Else
                EmitNode = "<m:groupChr><m:groupChrPr><m:chr m:val=""" & ModTexSymbols.EscapeXml(nd.Str1) & """ /><m:pos m:val=""" & nd.Str2 & """ /><m:vertJc m:val=""top"" /></m:groupChrPr><m:e>" & baseXml & "</m:e></m:groupChr>"
            End If

        Case AST_LIST
            res = ""
            For Each vItem In nd.ChildList
                res = res & EmitNode(CLng(vItem))
            Next vItem
            EmitNode = res

        Case Else
            EmitNode = ""
    End Select
End Function

Public Function TranspileLatexToOmml(ByVal rawLatex As String, ByVal isBlock As Boolean, ByRef outIsAmbiguous As Boolean, Optional ByRef outEqTag As String = "") As String
    Dim rootIdx As Long
    Dim innerXml As String

    On Error GoTo ErrHandler

    rootIdx = ModTexParser.ParseLatex(rawLatex, isBlock, rootIdx, outIsAmbiguous, outEqTag)
    If outIsAmbiguous Then
        TranspileLatexToOmml = ""
        Exit Function
    End If

    innerXml = EmitNode(rootIdx)
    TranspileLatexToOmml = "<m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">" & innerXml & "</m:oMath>"
    Exit Function

ErrHandler:
    outIsAmbiguous = True
    TranspileLatexToOmml = ""
End Function

Public Function WrapOmmlInFlatOpc(ByVal ommlXml As String) As String
    Dim pkg As String
    pkg = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
    pkg = pkg & "<?mso-application progid=""Word.Document""?>" & vbCrLf
    pkg = pkg & "<pkg:package xmlns:pkg=""http://schemas.microsoft.com/office/2006/xmlPackage"">" & vbCrLf
    pkg = pkg & "  <pkg:part pkg:name=""/_rels/.rels"" pkg:contentType=""application/vnd.openxmlformats-package.relationships+xml"">" & vbCrLf
    pkg = pkg & "    <pkg:xmlData>" & vbCrLf
    pkg = pkg & "      <Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" & vbCrLf
    pkg = pkg & "        <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>" & vbCrLf
    pkg = pkg & "      </Relationships>" & vbCrLf
    pkg = pkg & "    </pkg:xmlData>" & vbCrLf
    pkg = pkg & "  </pkg:part>" & vbCrLf
    pkg = pkg & "  <pkg:part pkg:name=""/word/document.xml"" pkg:contentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"">" & vbCrLf
    pkg = pkg & "    <pkg:xmlData>" & vbCrLf
    pkg = pkg & "      <w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">" & vbCrLf
    pkg = pkg & "        <w:body>" & vbCrLf
    pkg = pkg & "          <w:p>" & vbCrLf
    pkg = pkg & "            " & ommlXml & vbCrLf
    pkg = pkg & "          </w:p>" & vbCrLf
    pkg = pkg & "        </w:body>" & vbCrLf
    pkg = pkg & "      </w:document>" & vbCrLf
    pkg = pkg & "    </pkg:xmlData>" & vbCrLf
    pkg = pkg & "  </pkg:part>" & vbCrLf
    pkg = pkg & "</pkg:package>"
    WrapOmmlInFlatOpc = pkg
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
