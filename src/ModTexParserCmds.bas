Attribute VB_Name = "ModTexParserCmds"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexParserCmds
' Purpose: LaTeX command parser for fractions, radicals, fonts, nary, accents
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private Function IsEmptyAstNode(ByVal nd As Long) As Boolean
    If nd <= 0 Then
        IsEmptyAstNode = True
        Exit Function
    End If
    If ModTexAst.mNodes(nd).NodeType = AST_LIST Then
        If ModTexAst.mNodes(nd).ChildList Is Nothing Then
            IsEmptyAstNode = True
        ElseIf ModTexAst.mNodes(nd).ChildList.Count = 0 Then
            IsEmptyAstNode = True
        End If
    End If
End Function

Private Function IsSafeFontCmd(ByVal cmd As String, ByVal fontMode As String) As Boolean
    If cmd = "\ " Or cmd = "\," Or cmd = "\:" Or cmd = "\;" Or cmd = "\quad" Or cmd = "\qquad" Then
        IsSafeFontCmd = True
        Exit Function
    End If
    If cmd = "\%" Or cmd = "\$" Or cmd = "\&" Or cmd = "\_" Or cmd = "\#" Or cmd = "\{" Or cmd = "\}" Then
        IsSafeFontCmd = True
        Exit Function
    End If
    If fontMode = "text" Then
        If cmd = "\u" Or cmd = "\v" Or cmd = "\t" Or cmd = "\r" Or cmd = "\b" Or cmd = "\c" Or cmd = "\d" Or cmd = "\k" Or cmd = "\H" Then
            IsSafeFontCmd = True
            Exit Function
        End If
    End If
    If Len(ModTexSymbols.LookupGreek(cmd)) > 0 Then
        IsSafeFontCmd = True
        Exit Function
    End If
    If Len(ModTexSymbols.LookupSymbol(cmd)) > 0 Then
        IsSafeFontCmd = True
        Exit Function
    End If
    IsSafeFontCmd = False
End Function

Private Function HasComplexFontArg(ByVal fontMode As String) As Boolean
    Dim i As Long
    Dim depth As Long
    Dim t As TexToken

    t = ModTexAst.CurTok()
    If t.TokType = TOK_EOF Then Exit Function
    If t.TokType <> TOK_LBRACE Then
        If t.TokType = TOK_SUB Or t.TokType = TOK_SUP Then
            HasComplexFontArg = True
        ElseIf t.TokType = TOK_CMD Then
            HasComplexFontArg = Not IsSafeFontCmd(t.Val, fontMode)
        End If
        Exit Function
    End If

    depth = 1
    i = ModTexAst.mPos + 1
    Do While i <= ModTexAst.mTokenCount And depth > 0
        t = ModTexAst.mTokens(i)
        If t.TokType = TOK_EOF Then
            Exit Do
        ElseIf t.TokType = TOK_LBRACE Then
            depth = depth + 1
        ElseIf t.TokType = TOK_RBRACE Then
            depth = depth - 1
        ElseIf t.TokType = TOK_SUB Or t.TokType = TOK_SUP Then
            HasComplexFontArg = True
            Exit Function
        ElseIf t.TokType = TOK_CMD Then
            If Not IsSafeFontCmd(t.Val, fontMode) Then
                HasComplexFontArg = True
                Exit Function
            End If
        End If
        i = i + 1
    Loop
End Function

Private Function ParseFont(ByVal fontMode As String, Optional ByVal allowSubSup As Boolean = True) As Long
    Dim rawText As String
    Dim nd As Long
    ModTexAst.AdvanceTok
    If HasComplexFontArg(fontMode) Then
        ModTexAst.SetAmbiguous True
    End If
    rawText = ModTexAst.ConsumeLiteralTextArg()
    nd = ModTexAst.AllocNode(AST_TEXT)
    ModTexAst.mNodes(nd).Str1 = rawText
    ModTexAst.mNodes(nd).Str2 = fontMode
    If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
    ParseFont = nd
End Function

Private Function ConsumeLimitModOverride() As String
    Dim limOverride As String
    limOverride = ""
    Do While ModTexAst.CurTok().TokType = TOK_CMD
        If ModTexAst.CurTok().Val = "\limits" Then
            limOverride = "undOvr"
            ModTexAst.AdvanceTok
        ElseIf ModTexAst.CurTok().Val = "\nolimits" Then
            limOverride = "subSup"
            ModTexAst.AdvanceTok
        Else
            Exit Do
        End If
    Loop
    ConsumeLimitModOverride = limOverride
End Function

Private Function IsNaryStopTok() As Boolean
    Dim curT As TexToken
    Dim nextIdx As Long
    Dim nextVal As String

    curT = ModTexAst.CurTok()
    If curT.TokType = TOK_EOF Or curT.TokType = TOK_RBRACE Or curT.TokType = TOK_RPAREN Or curT.TokType = TOK_RBRACK Or curT.TokType = TOK_AMP Or curT.TokType = TOK_NEWLINE Then
        IsNaryStopTok = True
        Exit Function
    End If
    If curT.TokType = TOK_CMD Then
        If curT.Val = "\right" Or curT.Val = "\end" Then
            IsNaryStopTok = True
            Exit Function
        End If
        If ModTexSymbols.MapBigCloseCmd(curT.Val) <> "" Or ModTexSymbols.IsBigCloseOnlyCmd(curT.Val) Then
            nextVal = ModTexParserEnvs.PeekNextNonSpaceVal(ModTexAst.mPos, nextIdx)
            If ModTexParserEnvs.IsCloseDelim(nextVal) Then
                IsNaryStopTok = True
                Exit Function
            End If
        End If
    End If
    IsNaryStopTok = False
End Function

Private Function ParseNary(ByVal cmd As String, ByVal chVal As String) As Long
    Dim subNode As Long
    Dim supNode As Long
    Dim bodyNode As Long
    Dim nd As Long
    Dim p As Long
    Dim limOverride As String

    ModTexAst.AdvanceTok
    limOverride = ConsumeLimitModOverride()
    subNode = 0
    supNode = 0
    For p = 1 To 2
        If ModTexAst.CurTok().TokType = TOK_SUB And subNode = 0 Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParseArgument()
        ElseIf ModTexAst.CurTok().TokType = TOK_SUP And supNode = 0 Then
            ModTexAst.AdvanceTok
            supNode = ModTexParser.ParseArgument()
        End If
    Next p

    bodyNode = 0
    If Not IsNaryStopTok() Then
        bodyNode = ModTexParser.ParsePrimary()
    End If

    nd = ModTexAst.AllocNode(AST_NARY)
    ModTexAst.mNodes(nd).Str1 = chVal
    If Len(limOverride) > 0 Then
        ModTexAst.mNodes(nd).Str2 = limOverride
    ElseIf chVal = ChrW(&H222B) Or chVal = ChrW(&H222C) Or chVal = ChrW(&H222D) Or chVal = ChrW(&H222E) Then
        ModTexAst.mNodes(nd).Str2 = "subSup"
    Else
        ModTexAst.mNodes(nd).Str2 = "undOvr"
    End If
    ModTexAst.mNodes(nd).Child1 = subNode
    ModTexAst.mNodes(nd).Child2 = supNode
    ModTexAst.mNodes(nd).Child3 = bodyNode
    ParseNary = nd
End Function

Public Function ParseCommand(ByVal cmd As String, Optional ByVal allowSubSup As Boolean = True) As Long
    Dim cleanCmd As String, fontMode As String, begChr As String, endChr As String, rawDelim As String, chVal As String
    Dim numNode As Long, denNode As Long, baseNode As Long, degNode As Long, bodyNode As Long, subNode As Long, supNode As Long, fnNode As Long, nd As Long

    If cmd = "\tag" Or cmd = "\tag*" Then
        ModTexAst.AdvanceTok
        ModTexAst.mEqTag = ModTexAst.ConsumeLiteralTextArg()
        ParseCommand = 0
        Exit Function
    ElseIf cmd = "\label" Then
        ModTexAst.AdvanceTok
        rawDelim = ModTexAst.ConsumeLiteralTextArg()
        ParseCommand = 0
        Exit Function
    ElseIf ModTexSymbols.IsStyleMathCmd(cmd) Then
        Do While ModTexAst.CurTok().TokType = TOK_CMD
            If Not ModTexSymbols.IsStyleMathCmd(ModTexAst.CurTok().Val) Then Exit Do
            ModTexAst.AdvanceTok
        Loop
        If Not ModTexSymbols.IsExpressionStopTok(ModTexAst.CurTok().TokType, ModTexAst.CurTok().Val) Then
            If allowSubSup Then
                ParseCommand = ModTexParser.ParsePrimary()
            Else
                ParseCommand = ModTexParser.ParseArgument()
            End If
        End If
        Exit Function
    ElseIf ModTexSymbols.IsNoOpMathCmd(cmd) Then
        ModTexAst.AdvanceTok
        ParseCommand = 0
        Exit Function
    ElseIf cmd = "\frac" Or cmd = "\dfrac" Or cmd = "\tfrac" Or cmd = "\cfrac" Or cmd = "\binom" Or cmd = "\dbinom" Or cmd = "\tbinom" Then
        ModTexAst.AdvanceTok
        numNode = ModTexParser.ParseArgument()
        denNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_FRAC)
        ModTexAst.mNodes(nd).Child1 = numNode
        ModTexAst.mNodes(nd).Child2 = denNode
        If cmd = "\frac" Or cmd = "\dfrac" Or cmd = "\tfrac" Or cmd = "\cfrac" Then
            ModTexAst.mNodes(nd).Bool1 = True
            If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
            ParseCommand = nd
        Else
            ModTexAst.mNodes(nd).Bool1 = False
            baseNode = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(baseNode).Str1 = "("
            ModTexAst.mNodes(baseNode).Str2 = ")"
            ModTexAst.mNodes(baseNode).Bool1 = True
            ModTexAst.mNodes(baseNode).Child1 = nd
            If allowSubSup Then baseNode = ModTexAst.HandleSubSup(baseNode)
            ParseCommand = baseNode
        End If
        Exit Function
    ElseIf cmd = "\sqrt" Then
        ModTexAst.AdvanceTok
        degNode = 0
        If ModTexAst.CurTok().TokType = TOK_LBRACK Then
            ModTexAst.AdvanceTok
            degNode = ModTexParser.ParseExpression(TOK_RBRACK)
            ModTexAst.ExpectTok TOK_RBRACK
        End If
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_RAD)
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = degNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\left" Or ModTexSymbols.MapBigCloseCmd(cmd) <> "" Or ModTexSymbols.IsBigCloseOnlyCmd(cmd) Then
        ModTexAst.AdvanceTok
        If ModTexAst.CurTok().TokType = TOK_EOF Then
            ModTexAst.SetAmbiguous True
            ParseCommand = 0
            Exit Function
        End If
        rawDelim = ModTexAst.CurTok().Val
        begChr = ModTexParserEnvs.MapDelimToken(rawDelim)
        ModTexAst.AdvanceTok
        If cmd = "\left" Then
            bodyNode = ModTexParserEnvs.ParseDelimitedBlock("\right", rawDelim, endChr)
        ElseIf ModTexSymbols.MapBigCloseCmd(cmd) <> "" And ModTexParserEnvs.IsBigOpenMatched(cmd, rawDelim) Then
            bodyNode = ModTexParserEnvs.ParseDelimitedBlock(ModTexSymbols.MapBigCloseCmd(cmd), rawDelim, endChr)
        Else
            If Len(begChr) = 0 Then
                ParseCommand = 0
            Else
                nd = ModTexAst.AllocNode(AST_CHAR)
                ModTexAst.mNodes(nd).Str1 = begChr
                ModTexAst.mNodes(nd).Bool1 = False
                If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
                ParseCommand = nd
            End If
            Exit Function
        End If
        nd = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(nd).Str1 = begChr
        ModTexAst.mNodes(nd).Str2 = endChr
        ModTexAst.mNodes(nd).Bool1 = True
        ModTexAst.mNodes(nd).Child1 = bodyNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\begin" Then
        nd = ModTexParserEnvs.ParseEnvironment()
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\underbrace" Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_GROUPCHR)
        ModTexAst.mNodes(nd).Str1 = ChrW(&H23DF)
        ModTexAst.mNodes(nd).Str2 = "bot"
        ModTexAst.mNodes(nd).Child1 = baseNode
        subNode = 0
        If ModTexAst.CurTok().TokType = TOK_SUB Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParseArgument()
        End If
        If Not IsEmptyAstNode(subNode) Then
            baseNode = nd
            nd = ModTexAst.AllocNode(AST_LIMLOW)
            ModTexAst.mNodes(nd).Child1 = baseNode
            ModTexAst.mNodes(nd).Child2 = subNode
        End If
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\overbrace" Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_GROUPCHR)
        ModTexAst.mNodes(nd).Str1 = ChrW(&H23DE)
        ModTexAst.mNodes(nd).Str2 = "top"
        ModTexAst.mNodes(nd).Child1 = baseNode
        supNode = 0
        If ModTexAst.CurTok().TokType = TOK_SUP Then
            ModTexAst.AdvanceTok
            supNode = ModTexParser.ParseArgument()
        End If
        If Not IsEmptyAstNode(supNode) Then
            baseNode = nd
            nd = ModTexAst.AllocNode(AST_LIMUPP)
            ModTexAst.mNodes(nd).Child1 = baseNode
            ModTexAst.mNodes(nd).Child2 = supNode
        End If
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\overset" Then
        ModTexAst.AdvanceTok
        supNode = ModTexParser.ParseArgument()
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_LIMUPP)
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = supNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\underset" Then
        ModTexAst.AdvanceTok
        subNode = ModTexParser.ParseArgument()
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_LIMLOW)
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = subNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\xrightarrow" Or cmd = "\xleftarrow" Then
        ModTexAst.AdvanceTok
        subNode = 0
        If ModTexAst.CurTok().TokType = TOK_LBRACK Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParseExpression(TOK_RBRACK)
            ModTexAst.ExpectTok TOK_RBRACK
        End If
        supNode = ModTexParser.ParseArgument()
        baseNode = ModTexAst.AllocNode(AST_CHAR)
        If cmd = "\xrightarrow" Then
            ModTexAst.mNodes(baseNode).Str1 = ChrW(&H2192)
        Else
            ModTexAst.mNodes(baseNode).Str1 = ChrW(&H2190)
        End If
        ModTexAst.mNodes(baseNode).Bool1 = True
        If Not IsEmptyAstNode(subNode) Then
            nd = ModTexAst.AllocNode(AST_LIMLOW)
            ModTexAst.mNodes(nd).Child1 = baseNode
            ModTexAst.mNodes(nd).Child2 = subNode
            baseNode = nd
        End If
        If Not IsEmptyAstNode(supNode) Then
            nd = ModTexAst.AllocNode(AST_LIMUPP)
            ModTexAst.mNodes(nd).Child1 = baseNode
            ModTexAst.mNodes(nd).Child2 = supNode
            baseNode = nd
        End If
        If allowSubSup Then baseNode = ModTexAst.HandleSubSup(baseNode)
        ParseCommand = baseNode
        Exit Function
    ElseIf cmd = "\boxed" Then
        ModTexAst.AdvanceTok
        bodyNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_BORDERBOX)
        ModTexAst.mNodes(nd).Child1 = bodyNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\overline" Or cmd = "\underline" Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_BAR)
        If cmd = "\overline" Then
            ModTexAst.mNodes(nd).Str1 = "top"
        Else
            ModTexAst.mNodes(nd).Str1 = "bot"
        End If
        ModTexAst.mNodes(nd).Child1 = baseNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\overrightarrow" Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_GROUPCHR)
        ModTexAst.mNodes(nd).Str1 = ChrW(&H2192)
        ModTexAst.mNodes(nd).Str2 = "top"
        ModTexAst.mNodes(nd).Child1 = baseNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\operatorname*" Or cmd = "\operatorname" Then
        ModTexAst.AdvanceTok
        fnNode = ModTexAst.AllocNode(AST_TEXT)
        ModTexAst.mNodes(fnNode).Str1 = ModTexAst.ConsumeLiteralTextArg()
        ModTexAst.mNodes(fnNode).Str2 = "roman"
        ModTexAst.SkipLimitMods
        If allowSubSup And cmd = "\operatorname*" And ModTexAst.CurTok().TokType = TOK_SUB Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParseArgument()
            nd = ModTexAst.AllocNode(AST_LIMLOW)
            ModTexAst.mNodes(nd).Child1 = fnNode
            ModTexAst.mNodes(nd).Child2 = subNode
            nd = ModTexAst.HandleSubSup(nd)
            ParseCommand = nd
            Exit Function
        End If
        If allowSubSup Then fnNode = ModTexAst.HandleSubSup(fnNode)
        ParseCommand = fnNode
        Exit Function
    ElseIf cmd = "\not" Then
        ModTexAst.AdvanceTok
        chVal = ModTexSymbols.LookupNotNegation(ModTexAst.CurTok().Val)
        If Len(chVal) > 0 Then
            ModTexAst.AdvanceTok
            nd = ModTexAst.AllocNode(AST_CHAR)
            ModTexAst.mNodes(nd).Str1 = chVal
            ModTexAst.mNodes(nd).Bool1 = ModTexSymbols.IsOperatorChar(chVal)
            If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
            ParseCommand = nd
        Else
            ModTexAst.SetAmbiguous True
            ParseCommand = 0
        End If
        Exit Function
    End If

    fontMode = ModTexSymbols.LookupFontMode(cmd)
    If Len(fontMode) > 0 Then
        ParseCommand = ParseFont(fontMode, allowSubSup)
        Exit Function
    End If

    chVal = ModTexSymbols.LookupAccent(cmd)
    If Len(chVal) > 0 Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_ACC)
        ModTexAst.mNodes(nd).Str1 = chVal
        ModTexAst.mNodes(nd).Child1 = baseNode
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    End If

    cleanCmd = Mid$(cmd, 2)
    If ModTexSymbols.IsNamedFunction(cleanCmd) Then
        ModTexAst.AdvanceTok
        fnNode = ModTexAst.AllocNode(AST_TEXT)
        If cleanCmd = "argmin" Then
            ModTexAst.mNodes(fnNode).Str1 = "arg min"
        ElseIf cleanCmd = "argmax" Then
            ModTexAst.mNodes(fnNode).Str1 = "arg max"
        Else
            ModTexAst.mNodes(fnNode).Str1 = cleanCmd
        End If
        ModTexAst.mNodes(fnNode).Str2 = "roman"
        If cleanCmd = "lim" Or cleanCmd = "min" Or cleanCmd = "max" Or cleanCmd = "inf" Or cleanCmd = "sup" Or cleanCmd = "argmin" Or cleanCmd = "argmax" Then
            ModTexAst.SkipLimitMods
            If allowSubSup And ModTexAst.CurTok().TokType = TOK_SUB Then
                If ModTexAst.mIsBlock Or cleanCmd = "lim" Or cleanCmd = "argmin" Or cleanCmd = "argmax" Then
                    ModTexAst.AdvanceTok
                    subNode = ModTexParser.ParseArgument()
                    nd = ModTexAst.AllocNode(AST_LIMLOW)
                    ModTexAst.mNodes(nd).Child1 = fnNode
                    ModTexAst.mNodes(nd).Child2 = subNode
                    nd = ModTexAst.HandleSubSup(nd)
                    ParseCommand = nd
                    Exit Function
                End If
            End If
        End If
        If allowSubSup Then fnNode = ModTexAst.HandleSubSup(fnNode)
        ParseCommand = fnNode
        Exit Function
    End If

    chVal = ModTexSymbols.LookupNary(cmd)
    If Len(chVal) > 0 Then
        ParseCommand = ParseNary(cmd, chVal)
        Exit Function
    End If

    chVal = ModTexSymbols.LookupGreek(cmd)
    If Len(chVal) = 0 Then
        chVal = ModTexSymbols.LookupSymbol(cmd)
    End If
    If Len(chVal) > 0 Then
        ModTexAst.AdvanceTok
        nd = ModTexAst.AllocNode(AST_CHAR)
        ModTexAst.mNodes(nd).Str1 = chVal
        ModTexAst.mNodes(nd).Bool1 = ModTexSymbols.IsOperatorChar(chVal)
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    End If

    If cmd = "\|" Or cmd = "\Vert" Or cmd = "\parallel" Then
        ModTexAst.AdvanceTok
        nd = ModTexAst.AllocNode(AST_CHAR)
        ModTexAst.mNodes(nd).Str1 = ChrW(&H2225)
        ModTexAst.mNodes(nd).Bool1 = True
        If allowSubSup Then nd = ModTexAst.HandleSubSup(nd)
        ParseCommand = nd
        Exit Function
    End If

    chVal = ModTexSymbols.LookupSpace(cmd)
    If Len(chVal) > 0 Or cmd = "\!" Then
        ModTexAst.AdvanceTok
        If Len(chVal) > 0 Then
            nd = ModTexAst.AllocNode(AST_CHAR)
            ModTexAst.mNodes(nd).Str1 = chVal
            ModTexAst.mNodes(nd).Bool1 = False
            ParseCommand = nd
        Else
            ParseCommand = 0
        End If
        Exit Function
    End If

    ' Unknown command -> fail closed
    ModTexAst.AdvanceTok
    ModTexAst.SetAmbiguous True
    ParseCommand = 0
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
