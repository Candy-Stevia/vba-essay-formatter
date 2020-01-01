Attribute VB_Name = "ModTexParser"
Option Explicit

' ==============================================================================
' Module: ModTexParser
' Purpose: Primary, argument, and expression parser for LaTeX
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Function ParseArgument() As Long
    Dim tok As TexToken
    Dim nd As Long
    Dim ch As String

    If ModTexAst.CurTok().TokType = TOK_LBRACE Then
        ParseArgument = ModTexAst.ParseGroup()
        Exit Function
    End If

    tok = ModTexAst.CurTok()
    If tok.TokType = TOK_CMD Then
        ch = ModTexSymbols.LookupGreek(tok.Val)
        If Len(ch) = 0 Then
            ch = ModTexSymbols.LookupSymbol(tok.Val)
        End If
        If Len(ch) > 0 Then
            ModTexAst.AdvanceTok
            nd = ModTexAst.AllocNode(AST_CHAR)
            ModTexAst.mNodes(nd).Str1 = ch
            ModTexAst.mNodes(nd).Bool1 = ModTexSymbols.IsOperatorChar(ch)
            ParseArgument = nd
            Exit Function
        End If
        ParseArgument = ModTexParserCmds.ParseCommand(tok.Val)
        Exit Function
    End If

    If tok.TokType = TOK_CHAR Or tok.TokType = TOK_ESC_CHAR Or tok.TokType = TOK_LPAREN Or tok.TokType = TOK_RPAREN Or tok.TokType = TOK_LBRACK Or tok.TokType = TOK_RBRACK Then
        ModTexAst.AdvanceTok
        ch = tok.Val
        If ch = "-" Then
            ch = ChrW(&H2212)
        End If
        nd = ModTexAst.AllocNode(AST_CHAR)
        ModTexAst.mNodes(nd).Str1 = ch
        ModTexAst.mNodes(nd).Bool1 = ModTexSymbols.IsOperatorChar(ch)
        ParseArgument = nd
        Exit Function
    End If

    ParseArgument = ParsePrimary()
End Function

Public Function ParsePrimary() As Long
    Dim tok As TexToken
    Dim nd As Long
    Dim inner As Collection
    Dim p As Long
    Dim bodyNode As Long
    Dim matchIdx As Long
    Dim c As String
    Dim isOp As Boolean

    tok = ModTexAst.CurTok()

    If tok.TokType = TOK_CMD Then
        ParsePrimary = ModTexParserCmds.ParseCommand(tok.Val)
        Exit Function
    End If

    If tok.TokType = TOK_ESC_CHAR Then
        ModTexAst.AdvanceTok
        nd = ModTexAst.AllocNode(AST_CHAR)
        If tok.Val = " " Then
            ModTexAst.mNodes(nd).Str1 = ChrW(&H2005)
        Else
            ModTexAst.mNodes(nd).Str1 = tok.Val
        End If
        ModTexAst.mNodes(nd).Bool1 = False
        ParsePrimary = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    If tok.TokType = TOK_LPAREN Then
        matchIdx = ModTexAst.FindMatching(TOK_LPAREN, TOK_RPAREN)
        If matchIdx <> 0 Then
            ModTexAst.AdvanceTok
            Set inner = New Collection
            Do While ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RPAREN
                p = ParsePrimary()
                If p <> 0 Then
                    inner.Add p
                Else
                    If ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RPAREN Then
                        ModTexAst.AdvanceTok
                    End If
                End If
            Loop
            ModTexAst.ExpectTok TOK_RPAREN
            bodyNode = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(bodyNode).ChildList = inner
            nd = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(nd).Str1 = "("
            ModTexAst.mNodes(nd).Str2 = ")"
            ModTexAst.mNodes(nd).Bool1 = True
            ModTexAst.mNodes(nd).Child1 = bodyNode
            ParsePrimary = ModTexAst.HandleSubSup(nd)
            Exit Function
        End If
    End If

    If tok.TokType = TOK_LBRACK Then
        matchIdx = ModTexAst.FindMatching(TOK_LBRACK, TOK_RBRACK)
        If matchIdx <> 0 Then
            ModTexAst.AdvanceTok
            Set inner = New Collection
            Do While ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RBRACK
                p = ParsePrimary()
                If p <> 0 Then
                    inner.Add p
                Else
                    If ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RBRACK Then
                        ModTexAst.AdvanceTok
                    End If
                End If
            Loop
            ModTexAst.ExpectTok TOK_RBRACK
            bodyNode = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(bodyNode).ChildList = inner
            nd = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(nd).Str1 = "["
            ModTexAst.mNodes(nd).Str2 = "]"
            ModTexAst.mNodes(nd).Bool1 = True
            ModTexAst.mNodes(nd).Child1 = bodyNode
            ParsePrimary = ModTexAst.HandleSubSup(nd)
            Exit Function
        End If
    End If

    If tok.TokType = TOK_CHAR And tok.Val = "|" Then
        matchIdx = ModTexAst.FindMatchingBar()
        If matchIdx <> 0 Then
            ModTexAst.AdvanceTok
            Set inner = New Collection
            Do While ModTexAst.CurTok().TokType <> TOK_EOF
                If ModTexAst.CurTok().TokType = TOK_CHAR And ModTexAst.CurTok().Val = "|" Then
                    Exit Do
                End If
                p = ParsePrimary()
                If p <> 0 Then
                    inner.Add p
                Else
                    If ModTexAst.CurTok().TokType <> TOK_EOF And Not (ModTexAst.CurTok().TokType = TOK_CHAR And ModTexAst.CurTok().Val = "|") Then
                        ModTexAst.AdvanceTok
                    End If
                End If
            Loop
            ModTexAst.AdvanceTok
            bodyNode = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(bodyNode).ChildList = inner
            nd = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(nd).Str1 = "|"
            ModTexAst.mNodes(nd).Str2 = "|"
            ModTexAst.mNodes(nd).Bool1 = True
            ModTexAst.mNodes(nd).Child1 = bodyNode
            ParsePrimary = ModTexAst.HandleSubSup(nd)
            Exit Function
        End If
    End If

    If tok.TokType = TOK_CHAR Or tok.TokType = TOK_LPAREN Or tok.TokType = TOK_RPAREN Or tok.TokType = TOK_LBRACK Or tok.TokType = TOK_RBRACK Then
        ModTexAst.AdvanceTok
        c = tok.Val
        If c = "-" Then
            c = ChrW(&H2212)
        End If
        isOp = ModTexSymbols.IsOperatorChar(c)
        nd = ModTexAst.AllocNode(AST_CHAR)
        ModTexAst.mNodes(nd).Str1 = c
        ModTexAst.mNodes(nd).Bool1 = isOp
        ParsePrimary = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    ' Unknown or unhandled token -> advance to prevent infinite loop and flag ambiguous
    If ModTexAst.CurTok().TokType <> TOK_EOF Then
        ModTexAst.AdvanceTok
    End If
    ModTexAst.SetAmbiguous True
    ParsePrimary = 0
End Function

Public Function ParseExpression() As Long
    Dim items As Collection
    Dim p As Long
    Dim listNode As Long

    Set items = New Collection
    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        If ModTexAst.CurTok().TokType = TOK_RBRACE Or ModTexAst.CurTok().TokType = TOK_RBRACK Or ModTexAst.CurTok().TokType = TOK_RPAREN Or ModTexAst.CurTok().TokType = TOK_AMP Or ModTexAst.CurTok().TokType = TOK_NEWLINE Then
            Exit Do
        End If
        p = ParsePrimary()
        If p <> 0 Then
            items.Add p
        Else
            If ModTexAst.CurTok().TokType <> TOK_EOF Then
                ModTexAst.AdvanceTok
            End If
        End If
    Loop

    listNode = ModTexAst.AllocNode(AST_LIST)
    Set ModTexAst.mNodes(listNode).ChildList = items
    ParseExpression = listNode
End Function

Public Function ParseLatex(ByVal text As String, ByVal isBlock As Boolean, ByRef outRoot As Long, ByRef outIsAmbiguous As Boolean, ByRef outEqTag As String) As Long
    ModTexAst.ResetArena isBlock
    ModTexLexer.TokenizeLatex text
    outRoot = ParseExpression()
    outIsAmbiguous = ModTexAst.mIsAmbiguous
    outEqTag = ModTexAst.mEqTag
    ParseLatex = outRoot
End Function
