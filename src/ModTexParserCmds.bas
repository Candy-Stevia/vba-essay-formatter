Attribute VB_Name = "ModTexParserCmds"
Option Explicit

' ==============================================================================
' Module: ModTexParserCmds
' Purpose: LaTeX command parser for fractions, radicals, fonts, nary, accents
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private Function ParseFont(ByVal fontMode As String) As Long
    Dim rawText As String
    Dim nd As Long
    ModTexAst.AdvanceTok
    rawText = ModTexAst.ConsumeLiteralTextArg()
    nd = ModTexAst.AllocNode(AST_TEXT)
    ModTexAst.mNodes(nd).Str1 = rawText
    ModTexAst.mNodes(nd).Str2 = fontMode
    ParseFont = ModTexAst.HandleSubSup(nd)
End Function

Private Function ParseNary(ByVal cmd As String, ByVal chVal As String) As Long
    Dim subNode As Long
    Dim supNode As Long
    Dim bodyNode As Long
    Dim nd As Long
    Dim p As Long

    ModTexAst.AdvanceTok
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
    If ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RBRACE And ModTexAst.CurTok().TokType <> TOK_AMP And ModTexAst.CurTok().TokType <> TOK_NEWLINE Then
        bodyNode = ModTexParser.ParsePrimary()
    End If

    nd = ModTexAst.AllocNode(AST_NARY)
    ModTexAst.mNodes(nd).Str1 = chVal
    If chVal = ChrW(&H222B) Then
        ModTexAst.mNodes(nd).Str2 = "subSup"
    Else
        ModTexAst.mNodes(nd).Str2 = "undOvr"
    End If
    ModTexAst.mNodes(nd).Child1 = subNode
    ModTexAst.mNodes(nd).Child2 = supNode
    ModTexAst.mNodes(nd).Child3 = bodyNode
    ParseNary = nd
End Function

Public Function ParseCommand(ByVal cmd As String) As Long
    Dim cleanCmd As String
    Dim fontMode As String
    Dim numNode As Long
    Dim denNode As Long
    Dim baseNode As Long
    Dim degNode As Long
    Dim bodyNode As Long
    Dim subNode As Long
    Dim supNode As Long
    Dim fnNode As Long
    Dim nd As Long
    Dim begChr As String
    Dim endChr As String
    Dim chVal As String

    If cmd = "\tag" Then
        ModTexAst.AdvanceTok
        ModTexAst.mEqTag = ModTexAst.ConsumeLiteralTextArg()
        ParseCommand = 0
        Exit Function
    ElseIf cmd = "\displaystyle" Then
        ModTexAst.AdvanceTok
        ParseCommand = ModTexParser.ParsePrimary()
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
            ParseCommand = ModTexAst.HandleSubSup(nd)
        Else
            ModTexAst.mNodes(nd).Bool1 = False
            baseNode = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(baseNode).Str1 = "("
            ModTexAst.mNodes(baseNode).Str2 = ")"
            ModTexAst.mNodes(baseNode).Bool1 = True
            ModTexAst.mNodes(baseNode).Child1 = nd
            ParseCommand = ModTexAst.HandleSubSup(baseNode)
        End If
        Exit Function
    ElseIf cmd = "\sqrt" Then
        ModTexAst.AdvanceTok
        degNode = 0
        If ModTexAst.CurTok().TokType = TOK_LBRACK Then
            ModTexAst.AdvanceTok
            degNode = ModTexParser.ParsePrimary()
            ModTexAst.ExpectTok TOK_RBRACK
        End If
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_RAD)
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = degNode
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    ElseIf cmd = "\left" Or cmd = "\bigg" Or cmd = "\Big" Or cmd = "\Bigg" Then
        ModTexAst.AdvanceTok
        begChr = ModTexParserEnvs.MapDelimToken(ModTexAst.CurTok().Val)
        ModTexAst.AdvanceTok
        If cmd = "\left" Then
            bodyNode = ModTexParserEnvs.ParseDelimitedBlock("\right", endChr)
        Else
            bodyNode = ModTexParserEnvs.ParseDelimitedBlock(cmd, endChr)
        End If
        nd = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(nd).Str1 = begChr
        ModTexAst.mNodes(nd).Str2 = endChr
        ModTexAst.mNodes(nd).Bool1 = True
        ModTexAst.mNodes(nd).Child1 = bodyNode
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    ElseIf cmd = "\begin" Then
        nd = ModTexParserEnvs.ParseEnvironment()
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    ElseIf cmd = "\underbrace" Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        subNode = 0
        If ModTexAst.CurTok().TokType = TOK_SUB Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParseArgument()
        End If
        nd = ModTexAst.AllocNode(AST_LIMLOW)
        ModTexAst.mNodes(nd).Str1 = "underbrace"
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = subNode
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\overset" Then
        ModTexAst.AdvanceTok
        supNode = ModTexParser.ParseArgument()
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_LIMUPP)
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = supNode
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\xrightarrow" Then
        ModTexAst.AdvanceTok
        subNode = 0
        If ModTexAst.CurTok().TokType = TOK_LBRACK Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParsePrimary()
            ModTexAst.ExpectTok TOK_RBRACK
        End If
        supNode = ModTexParser.ParseArgument()
        baseNode = ModTexAst.AllocNode(AST_CHAR)
        ModTexAst.mNodes(baseNode).Str1 = ChrW(&H2192)
        nd = ModTexAst.AllocNode(AST_LIMUPP)
        ModTexAst.mNodes(nd).Child1 = baseNode
        ModTexAst.mNodes(nd).Child2 = supNode
        ParseCommand = nd
        Exit Function
    ElseIf cmd = "\boxed" Then
        ModTexAst.AdvanceTok
        bodyNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_BORDERBOX)
        ModTexAst.mNodes(nd).Child1 = bodyNode
        ParseCommand = ModTexAst.HandleSubSup(nd)
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
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    ElseIf cmd = "\overrightarrow" Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_GROUPCHR)
        ModTexAst.mNodes(nd).Str1 = ChrW(&H2192)
        ModTexAst.mNodes(nd).Str2 = "top"
        ModTexAst.mNodes(nd).Child1 = baseNode
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    ElseIf cmd = "\operatorname*" Or cmd = "\operatorname" Then
        ModTexAst.AdvanceTok
        fnNode = ModTexAst.AllocNode(AST_TEXT)
        ModTexAst.mNodes(fnNode).Str1 = ModTexAst.ConsumeLiteralTextArg()
        ModTexAst.mNodes(fnNode).Str2 = "roman"
        If cmd = "\operatorname*" And ModTexAst.CurTok().TokType = TOK_SUB Then
            ModTexAst.AdvanceTok
            subNode = ModTexParser.ParseArgument()
            nd = ModTexAst.AllocNode(AST_LIMLOW)
            ModTexAst.mNodes(nd).Child1 = fnNode
            ModTexAst.mNodes(nd).Child2 = subNode
            ParseCommand = nd
            Exit Function
        End If
        ParseCommand = ModTexAst.HandleSubSup(fnNode)
        Exit Function
    ElseIf cmd = "\not" Then
        ModTexAst.AdvanceTok
        chVal = ModTexSymbols.LookupNotNegation(ModTexAst.CurTok().Val)
        If Len(chVal) > 0 Then
            ModTexAst.AdvanceTok
            nd = ModTexAst.AllocNode(AST_CHAR)
            ModTexAst.mNodes(nd).Str1 = chVal
            ModTexAst.mNodes(nd).Bool1 = ModTexSymbols.IsOperatorChar(chVal)
            ParseCommand = ModTexAst.HandleSubSup(nd)
        Else
            ModTexAst.SetAmbiguous True
            ParseCommand = 0
        End If
        Exit Function
    End If

    fontMode = ModTexSymbols.LookupFontMode(cmd)
    If Len(fontMode) > 0 Then
        ParseCommand = ParseFont(fontMode)
        Exit Function
    End If

    chVal = ModTexSymbols.LookupAccent(cmd)
    If Len(chVal) > 0 Then
        ModTexAst.AdvanceTok
        baseNode = ModTexParser.ParseArgument()
        nd = ModTexAst.AllocNode(AST_ACC)
        ModTexAst.mNodes(nd).Str1 = chVal
        ModTexAst.mNodes(nd).Child1 = baseNode
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    cleanCmd = Mid$(cmd, 2)
    If ModTexSymbols.IsNamedFunction(cleanCmd) Then
        ModTexAst.AdvanceTok
        fnNode = ModTexAst.AllocNode(AST_TEXT)
        If cleanCmd = "argmin" Then
            ModTexAst.mNodes(fnNode).Str1 = "arg min"
        Else
            ModTexAst.mNodes(fnNode).Str1 = cleanCmd
        End If
        ModTexAst.mNodes(fnNode).Str2 = "roman"
        If cleanCmd = "lim" Or cleanCmd = "min" Or cleanCmd = "max" Or cleanCmd = "inf" Or cleanCmd = "sup" Or cleanCmd = "argmin" Then
            If ModTexAst.CurTok().TokType = TOK_SUB Then
                ModTexAst.AdvanceTok
                subNode = ModTexParser.ParseArgument()
                If Not ModTexAst.mIsBlock And cleanCmd <> "lim" And cleanCmd <> "argmin" Then
                    nd = ModTexAst.AllocNode(AST_SUBSUP)
                    ModTexAst.mNodes(nd).Child1 = fnNode
                    ModTexAst.mNodes(nd).Child2 = subNode
                    ParseCommand = nd
                    Exit Function
                Else
                    nd = ModTexAst.AllocNode(AST_LIMLOW)
                    ModTexAst.mNodes(nd).Child1 = fnNode
                    ModTexAst.mNodes(nd).Child2 = subNode
                    ParseCommand = nd
                    Exit Function
                End If
            End If
        End If
        ParseCommand = ModTexAst.HandleSubSup(fnNode)
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
        ParseCommand = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    If cmd = "\|" Or cmd = "\Vert" Or cmd = "\parallel" Then
        ModTexAst.AdvanceTok
        nd = ModTexAst.AllocNode(AST_CHAR)
        ModTexAst.mNodes(nd).Str1 = ChrW(&H2225)
        ModTexAst.mNodes(nd).Bool1 = True
        ParseCommand = ModTexAst.HandleSubSup(nd)
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
