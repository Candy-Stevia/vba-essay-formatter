Attribute VB_Name = "ModTexParserEnvs"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexParserEnvs
' Purpose: Environment parser for cases, aligned, align*, pmatrix, bmatrix, array
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Function ParseEnvironment() As Long
    ModTexAst.mDepth = ModTexAst.mDepth + 1
    If ModTexAst.mDepth > ModTexAst.MAX_DEPTH Then
        ModTexAst.SetAmbiguous True
        ModTexAst.mDepth = ModTexAst.mDepth - 1
        ParseEnvironment = 0
        Exit Function
    End If
    ParseEnvironment = ParseEnvironmentCore()
    ModTexAst.mDepth = ModTexAst.mDepth - 1
End Function

Private Function ParseEnvironmentCore() As Long
    Dim envName As String, endName As String, colSpec As String, colAligns As String, ch As String, alName As String
    Dim i As Long, cellList As Long, pNode As Long, matNode As Long, prevPos As Long
    Dim rows As Collection, curRow As Collection, cellItems As Collection
    Dim curT As TexToken
    Dim foundEnd As Boolean

    ' Expect \begin
    ModTexAst.AdvanceTok
    envName = Trim$(ModTexAst.ConsumeLiteralTextArg())

    ' Determine column alignments
    If envName = "cases" Then
        colAligns = "left,left"
    ElseIf envName = "aligned" Or envName = "align*" Or envName = "align" Or envName = "split" Then
        colAligns = "right,left,right,left"
    ElseIf envName = "gathered" Or envName = "multline" Or envName = "multline*" Or envName = "equation" Or envName = "equation*" Then
        colAligns = "center"
    ElseIf envName = "pmatrix" Or envName = "bmatrix" Or envName = "vmatrix" Or envName = "Vmatrix" Or envName = "Bmatrix" Or envName = "matrix" Or envName = "smallmatrix" Then
        colAligns = "center,center,center,center"
    ElseIf envName = "array" Then
        colSpec = ModTexAst.ConsumeLiteralTextArg()
        If InStr(1, colSpec, "|") > 0 Or InStr(1, colSpec, "\hline") > 0 Then
            ModTexAst.SetAmbiguous True
            ParseEnvironmentCore = 0
            Exit Function
        End If
        colAligns = ""
        For i = 1 To Len(colSpec)
            ch = Mid$(colSpec, i, 1)
            alName = ""
            If ch = "l" Then
                alName = "left"
            ElseIf ch = "r" Then
                alName = "right"
            ElseIf ch = "c" Then
                alName = "center"
            End If
            If Len(alName) > 0 Then
                If Len(colAligns) > 0 Then
                    colAligns = colAligns & ","
                End If
                colAligns = colAligns & alName
            End If
        Next i
        If Len(colAligns) = 0 Then
            colAligns = "center"
        End If
    Else
        ' Unsupported environment -> fail closed
        ModTexAst.SetAmbiguous True
        ParseEnvironmentCore = 0
        Exit Function
    End If

    Set rows = New Collection
    Set curRow = New Collection
    Set cellItems = New Collection
    foundEnd = False

    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        curT = ModTexAst.CurTok()

        If curT.TokType = TOK_CMD Then
            If curT.Val = "\end" Then
                ModTexAst.AdvanceTok
                endName = Trim$(ModTexAst.ConsumeLiteralTextArg())
                If endName <> envName Then
                    ModTexAst.SetAmbiguous True
                    ParseEnvironmentCore = 0
                    Exit Function
                End If
                cellList = ModTexAst.AllocNode(AST_LIST)
                Set ModTexAst.mNodes(cellList).ChildList = cellItems
                If envName = "equation" Or envName = "equation*" Then
                    ParseEnvironmentCore = cellList
                    Exit Function
                End If
                curRow.Add cellList
                rows.Add curRow
                foundEnd = True
                Exit Do
            ElseIf curT.Val = "\hline" Then
                ModTexAst.SetAmbiguous True
                ParseEnvironmentCore = 0
                Exit Function
            End If
        End If

        If curT.TokType = TOK_AMP Then
            ModTexAst.AdvanceTok
            cellList = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(cellList).ChildList = cellItems
            curRow.Add cellList
            Set cellItems = New Collection
        ElseIf curT.TokType = TOK_NEWLINE Then
            ModTexAst.AdvanceTok
            cellList = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(cellList).ChildList = cellItems
            curRow.Add cellList
            rows.Add curRow
            Set curRow = New Collection
            Set cellItems = New Collection
        Else
            prevPos = ModTexAst.mPos
            pNode = ModTexParser.ParsePrimary()
            If pNode <> 0 Then
                cellItems.Add pNode
            Else
                If ModTexAst.mPos = prevPos And ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_AMP And ModTexAst.CurTok().TokType <> TOK_NEWLINE Then
                    ModTexAst.AdvanceTok
                End If
            End If
        End If
    Loop

    If Not foundEnd Then
        ModTexAst.SetAmbiguous True
        ParseEnvironmentCore = 0
        Exit Function
    End If

    matNode = ModTexAst.AllocNode(AST_MATRIX)
    ModTexAst.mNodes(matNode).Str1 = colAligns
    Set ModTexAst.mNodes(matNode).ChildList = rows

    If envName = "cases" Then
        ParseEnvironmentCore = WrapMatrixDelim(matNode, "{", "")
    ElseIf envName = "pmatrix" Then
        ParseEnvironmentCore = WrapMatrixDelim(matNode, "(", ")")
    ElseIf envName = "bmatrix" Then
        ParseEnvironmentCore = WrapMatrixDelim(matNode, "[", "]")
    ElseIf envName = "vmatrix" Then
        ParseEnvironmentCore = WrapMatrixDelim(matNode, "|", "|")
    ElseIf envName = "Vmatrix" Then
        ParseEnvironmentCore = WrapMatrixDelim(matNode, ChrW(&H2225), ChrW(&H2225))
    ElseIf envName = "Bmatrix" Then
        ParseEnvironmentCore = WrapMatrixDelim(matNode, "{", "}")
    Else
        ParseEnvironmentCore = matNode
    End If
End Function

Private Function WrapMatrixDelim(ByVal matNode As Long, ByVal bChr As String, ByVal eChr As String) As Long
    Dim delimNode As Long
    delimNode = ModTexAst.AllocNode(AST_DELIM)
    ModTexAst.mNodes(delimNode).Str1 = bChr
    ModTexAst.mNodes(delimNode).Str2 = eChr
    ModTexAst.mNodes(delimNode).Bool1 = True
    ModTexAst.mNodes(delimNode).Child1 = matNode
    WrapMatrixDelim = delimNode
End Function

Public Function MapDelimToken(ByVal rawDelim As String) As String
    Select Case rawDelim
        Case "\{", "{", "\lbrace"
            MapDelimToken = "{"
        Case "\}", "}", "\rbrace"
            MapDelimToken = "}"
        Case "\lbrack"
            MapDelimToken = "["
        Case "\rbrack"
            MapDelimToken = "]"
        Case "|", "\vert"
            MapDelimToken = "|"
        Case "\backslash"
            MapDelimToken = "\"
        Case "\langle"
            MapDelimToken = ChrW(&H27E8)
        Case "\rangle"
            MapDelimToken = ChrW(&H27E9)
        Case "\lfloor"
            MapDelimToken = ChrW(&H230A)
        Case "\rfloor"
            MapDelimToken = ChrW(&H230B)
        Case "\lceil"
            MapDelimToken = ChrW(&H2308)
        Case "\rceil"
            MapDelimToken = ChrW(&H2309)
        Case "\|", "\Vert"
            MapDelimToken = ChrW(&H2225)
        Case "."
            MapDelimToken = ""
        Case Else
            If Len(rawDelim) = 1 Then
                MapDelimToken = rawDelim
            Else
                ModTexAst.SetAmbiguous True
                MapDelimToken = ""
            End If
    End Select
End Function

Public Function IsOpenDelim(ByVal s As String) As Boolean
    IsOpenDelim = (InStr(1, "|(|[|\{|{|\lbrace|\lbrack|\langle|\lfloor|\lceil|", "|" & s & "|") > 0)
End Function

Public Function IsCloseDelim(ByVal s As String) As Boolean
    IsCloseDelim = (InStr(1, "|)|]|\}|}|\rbrace|\rbrack|\rangle|\rfloor|\rceil|.|", "|" & s & "|") > 0)
End Function

Public Function IsSymmetricDelim(ByVal s As String) As Boolean
    IsSymmetricDelim = (InStr(1, "|||\vert|\||\Vert|", "|" & s & "|") > 0)
End Function

Public Function PeekNextNonSpaceVal(ByVal afterIdx As Long, ByRef outIdx As Long) As String
    Dim i As Long
    i = afterIdx + 1
    Do While i <= ModTexAst.mTokenCount
        If ModTexAst.mTokens(i).TokType <> TOK_SPACE Then
            outIdx = i
            PeekNextNonSpaceVal = ModTexAst.mTokens(i).Val
            Exit Function
        End If
        i = i + 1
    Loop
    outIdx = ModTexAst.mTokenCount
    PeekNextNonSpaceVal = ""
End Function

Public Function IsBigOpenMatched(ByVal cmd As String, ByVal rawDelim As String) As Boolean
    Dim i As Long, nextIdx As Long, dParen As Long, dBrack As Long, dBrace As Long, dLeft As Long, dEnv As Long, dBig As Long
    Dim t As TexToken
    Dim nextVal As String
    Dim isSym As Boolean, atTop As Boolean, closeCmd As String

    closeCmd = ModTexSymbols.MapBigCloseCmd(cmd)
    isSym = IsSymmetricDelim(rawDelim)
    If Not isSym And Not IsOpenDelim(rawDelim) Then
        IsBigOpenMatched = False
        Exit Function
    End If

    i = ModTexAst.mPos
    Do While i <= ModTexAst.mTokenCount
        t = ModTexAst.mTokens(i)
        atTop = (dParen = 0 And dBrack = 0 And dBrace = 0 And dLeft = 0 And dEnv = 0)
        If t.TokType = TOK_EOF Then
            Exit Do
        ElseIf t.TokType = TOK_LPAREN Then
            dParen = dParen + 1
        ElseIf t.TokType = TOK_RPAREN Then
            If dParen = 0 Then
                Exit Do
            End If
            dParen = dParen - 1
        ElseIf t.TokType = TOK_LBRACK Then
            dBrack = dBrack + 1
        ElseIf t.TokType = TOK_RBRACK Then
            If dBrack = 0 Then
                Exit Do
            End If
            dBrack = dBrack - 1
        ElseIf t.TokType = TOK_LBRACE Then
            dBrace = dBrace + 1
        ElseIf t.TokType = TOK_RBRACE Then
            If dBrace = 0 Then
                Exit Do
            End If
            dBrace = dBrace - 1
        ElseIf (t.TokType = TOK_AMP Or t.TokType = TOK_NEWLINE) And dEnv = 0 Then
            Exit Do
        ElseIf t.TokType = TOK_CMD Then
            If t.Val = "\left" Then
                dLeft = dLeft + 1
                nextVal = PeekNextNonSpaceVal(i, nextIdx)
                i = nextIdx
            ElseIf t.Val = "\right" Then
                If dLeft = 0 Then
                    Exit Do
                End If
                dLeft = dLeft - 1
                nextVal = PeekNextNonSpaceVal(i, nextIdx)
                i = nextIdx
            ElseIf t.Val = "\begin" Then
                dEnv = dEnv + 1
            ElseIf t.Val = "\end" Then
                If dEnv = 0 Then
                    Exit Do
                End If
                dEnv = dEnv - 1
            ElseIf (t.Val = cmd Or t.Val = closeCmd) And atTop Then
                nextVal = PeekNextNonSpaceVal(i, nextIdx)
                If isSym And IsSymmetricDelim(nextVal) Then
                    If t.Val = closeCmd And MapDelimToken(nextVal) = MapDelimToken(rawDelim) And dBig = 0 Then
                        IsBigOpenMatched = True
                        Exit Function
                    End If
                ElseIf t.Val = cmd And IsOpenDelim(nextVal) Then
                    dBig = dBig + 1
                ElseIf t.Val = closeCmd And IsCloseDelim(nextVal) Then
                    If dBig = 0 Then
                        If Not isSym Then
                            IsBigOpenMatched = True
                            Exit Function
                        End If
                        Exit Do
                    End If
                    dBig = dBig - 1
                End If
                i = nextIdx
            End If
        End If
        i = i + 1
    Loop
    IsBigOpenMatched = False
End Function

Public Function ParseDelimitedBlock(ByVal endCmd As String, ByVal begRaw As String, ByRef outEndChr As String) As Long
    ModTexAst.mDepth = ModTexAst.mDepth + 1
    If ModTexAst.mDepth > ModTexAst.MAX_DEPTH Then
        ModTexAst.SetAmbiguous True
        ModTexAst.mDepth = ModTexAst.mDepth - 1
        outEndChr = ""
        ParseDelimitedBlock = 0
        Exit Function
    End If
    ParseDelimitedBlock = ParseDelimitedBlockCore(endCmd, begRaw, outEndChr)
    ModTexAst.mDepth = ModTexAst.mDepth - 1
End Function

Private Function ParseDelimitedBlockCore(ByVal endCmd As String, ByVal begRaw As String, ByRef outEndChr As String) As Long
    Dim inner As Collection
    Dim p As Long, prevPos As Long, bodyNode As Long, nextIdx As Long
    Dim foundEnd As Boolean, isMatch As Boolean, isSym As Boolean
    Dim nextVal As String

    Set inner = New Collection
    foundEnd = False
    isSym = IsSymmetricDelim(begRaw)
    outEndChr = ""

    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        If ModTexAst.CurTok().TokType = TOK_CMD And ModTexAst.CurTok().Val = endCmd Then
            isMatch = False
            If endCmd = "\right" Then
                isMatch = True
            Else
                nextVal = PeekNextNonSpaceVal(ModTexAst.mPos, nextIdx)
                If isSym Then
                    isMatch = IsSymmetricDelim(nextVal) And (MapDelimToken(nextVal) = MapDelimToken(begRaw))
                Else
                    isMatch = IsCloseDelim(nextVal)
                End If
            End If
            If isMatch Then
                ModTexAst.AdvanceTok
                If ModTexAst.CurTok().TokType = TOK_EOF Then
                    ModTexAst.SetAmbiguous True
                Else
                    outEndChr = MapDelimToken(ModTexAst.CurTok().Val)
                    ModTexAst.AdvanceTok
                End If
                foundEnd = True
                Exit Do
            End If
        End If
        prevPos = ModTexAst.mPos
        p = ModTexParser.ParsePrimary()
        If p <> 0 Then
            inner.Add p
        ElseIf ModTexAst.mPos = prevPos And ModTexAst.CurTok().TokType <> TOK_EOF Then
            ModTexAst.AdvanceTok
        End If
    Loop

    If Not foundEnd Then
        ModTexAst.SetAmbiguous True
    End If
    bodyNode = ModTexAst.AllocNode(AST_LIST)
    Set ModTexAst.mNodes(bodyNode).ChildList = inner
    ParseDelimitedBlockCore = bodyNode
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
