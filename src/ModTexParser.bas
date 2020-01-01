Attribute VB_Name = "ModTexParser"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexParser
' Purpose: Primary, argument, and expression parser for LaTeX
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Function ParseArgument() As Long
    ModTexAst.mDepth = ModTexAst.mDepth + 1
    If ModTexAst.mDepth > ModTexAst.MAX_DEPTH Then
        ModTexAst.SetAmbiguous True
        ModTexAst.mDepth = ModTexAst.mDepth - 1
        ParseArgument = 0
        Exit Function
    End If
    ParseArgument = ParseArgumentCore()
    ModTexAst.mDepth = ModTexAst.mDepth - 1
End Function

Private Function ParseArgumentCore() As Long
    Dim tok As TexToken
    Dim nd As Long
    Dim ch As String

    If ModTexAst.CurTok().TokType = TOK_LBRACE Then
        ParseArgumentCore = ModTexAst.ParseGroup()
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
            ParseArgumentCore = nd
            Exit Function
        End If
        ParseArgumentCore = ModTexParserCmds.ParseCommand(tok.Val)
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
        ParseArgumentCore = nd
        Exit Function
    End If

    ParseArgumentCore = ParsePrimary()
End Function

Public Function ParseGroupCore() As Long
    Dim items As Collection
    Dim pNode As Long
    Dim listNode As Long
    Dim prevPos As Long

    ModTexAst.mDepth = ModTexAst.mDepth + 1
    If ModTexAst.mDepth > ModTexAst.MAX_DEPTH Then
        ModTexAst.SetAmbiguous True
        ModTexAst.mDepth = ModTexAst.mDepth - 1
        ParseGroupCore = 0
        Exit Function
    End If

    ModTexAst.ExpectTok TOK_LBRACE
    Set items = New Collection

    Do While ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RBRACE
        prevPos = ModTexAst.mPos
        pNode = ParsePrimary()
        If pNode <> 0 Then
            items.Add pNode
        Else
            If ModTexAst.mPos = prevPos And ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_RBRACE Then
                ModTexAst.AdvanceTok
            End If
        End If
    Loop

    ModTexAst.ExpectTok TOK_RBRACE
    listNode = ModTexAst.AllocNode(AST_LIST)
    Set ModTexAst.mNodes(listNode).ChildList = items
    ModTexAst.mDepth = ModTexAst.mDepth - 1
    ParseGroupCore = listNode
End Function

Private Function ParseDelimInner(ByVal closeType As Integer, ByVal closeVal As String) As Long
    Dim inner As Collection
    Dim p As Long
    Dim prevPos As Long
    Dim bodyNode As Long
    Dim curT As TexToken

    Set inner = New Collection
    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        curT = ModTexAst.CurTok()
        If curT.TokType = closeType Then
            If Len(closeVal) = 0 Or curT.Val = closeVal Then
                Exit Do
            End If
        End If
        prevPos = ModTexAst.mPos
        p = ParsePrimary()
        If p <> 0 Then
            inner.Add p
        ElseIf ModTexAst.mPos = prevPos And ModTexAst.CurTok().TokType <> TOK_EOF Then
            ModTexAst.AdvanceTok
        End If
    Loop
    bodyNode = ModTexAst.AllocNode(AST_LIST)
    Set ModTexAst.mNodes(bodyNode).ChildList = inner
    ParseDelimInner = bodyNode
End Function

Public Function ParsePrimary() As Long
    ModTexAst.mDepth = ModTexAst.mDepth + 1
    If ModTexAst.mDepth > ModTexAst.MAX_DEPTH Then
        ModTexAst.SetAmbiguous True
        ModTexAst.mDepth = ModTexAst.mDepth - 1
        ParsePrimary = 0
        Exit Function
    End If
    ParsePrimary = ParsePrimaryCore()
    ModTexAst.mDepth = ModTexAst.mDepth - 1
End Function

Private Function ParsePrimaryCore() As Long
    Dim tok As TexToken
    Dim nd As Long
    Dim bodyNode As Long
    Dim matchIdx As Long
    Dim c As String
    Dim isOp As Boolean

    tok = ModTexAst.CurTok()

    If tok.TokType = TOK_LBRACE Then
        nd = ModTexAst.ParseGroup()
        ParsePrimaryCore = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    If tok.TokType = TOK_CMD Then
        ParsePrimaryCore = ModTexParserCmds.ParseCommand(tok.Val)
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
        ParsePrimaryCore = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    If tok.TokType = TOK_LPAREN Then
        matchIdx = ModTexAst.FindMatching(TOK_LPAREN, TOK_RPAREN)
        If matchIdx <> 0 Then
            ModTexAst.AdvanceTok
            bodyNode = ParseDelimInner(TOK_RPAREN, "")
            ModTexAst.ExpectTok TOK_RPAREN
            nd = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(nd).Str1 = "("
            ModTexAst.mNodes(nd).Str2 = ")"
            ModTexAst.mNodes(nd).Bool1 = True
            ModTexAst.mNodes(nd).Child1 = bodyNode
            ParsePrimaryCore = ModTexAst.HandleSubSup(nd)
            Exit Function
        End If
    End If

    If tok.TokType = TOK_LBRACK Then
        matchIdx = ModTexAst.FindMatching(TOK_LBRACK, TOK_RBRACK)
        If matchIdx <> 0 Then
            ModTexAst.AdvanceTok
            bodyNode = ParseDelimInner(TOK_RBRACK, "")
            ModTexAst.ExpectTok TOK_RBRACK
            nd = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(nd).Str1 = "["
            ModTexAst.mNodes(nd).Str2 = "]"
            ModTexAst.mNodes(nd).Bool1 = True
            ModTexAst.mNodes(nd).Child1 = bodyNode
            ParsePrimaryCore = ModTexAst.HandleSubSup(nd)
            Exit Function
        End If
    End If

    If tok.TokType = TOK_CHAR And tok.Val = "|" Then
        matchIdx = ModTexAst.FindMatchingBar()
        If matchIdx <> 0 Then
            ModTexAst.AdvanceTok
            bodyNode = ParseDelimInner(TOK_CHAR, "|")
            ModTexAst.AdvanceTok
            nd = ModTexAst.AllocNode(AST_DELIM)
            ModTexAst.mNodes(nd).Str1 = "|"
            ModTexAst.mNodes(nd).Str2 = "|"
            ModTexAst.mNodes(nd).Bool1 = True
            ModTexAst.mNodes(nd).Child1 = bodyNode
            ParsePrimaryCore = ModTexAst.HandleSubSup(nd)
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
        ParsePrimaryCore = ModTexAst.HandleSubSup(nd)
        Exit Function
    End If

    ' Unknown or unhandled token -> advance to prevent infinite loop and flag ambiguous
    If ModTexAst.CurTok().TokType <> TOK_EOF Then
        ModTexAst.AdvanceTok
    End If
    ModTexAst.SetAmbiguous True
    ParsePrimaryCore = 0
End Function

Public Function ParseExpression(Optional ByVal stopTokType As Integer = 0) As Long
    ModTexAst.mDepth = ModTexAst.mDepth + 1
    If ModTexAst.mDepth > ModTexAst.MAX_DEPTH Then
        ModTexAst.SetAmbiguous True
        ModTexAst.mDepth = ModTexAst.mDepth - 1
        ParseExpression = 0
        Exit Function
    End If
    ParseExpression = ParseExpressionCore(stopTokType)
    ModTexAst.mDepth = ModTexAst.mDepth - 1
End Function

Private Function ParseExpressionCore(ByVal stopTokType As Integer) As Long
    Dim items As Collection
    Dim p As Long
    Dim prevPos As Long
    Dim listNode As Long
    Dim curT As TexToken

    Set items = New Collection
    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        curT = ModTexAst.CurTok()
        If stopTokType <> 0 And curT.TokType = stopTokType Then
            Exit Do
        End If
        If curT.TokType = TOK_RBRACE Or curT.TokType = TOK_AMP Or curT.TokType = TOK_NEWLINE Then
            Exit Do
        End If
        prevPos = ModTexAst.mPos
        p = ParsePrimary()
        If p <> 0 Then
            items.Add p
        Else
            If ModTexAst.mPos = prevPos And ModTexAst.CurTok().TokType <> TOK_EOF Then
                ModTexAst.AdvanceTok
            End If
        End If
    Loop

    listNode = ModTexAst.AllocNode(AST_LIST)
    Set ModTexAst.mNodes(listNode).ChildList = items
    ParseExpressionCore = listNode
End Function

Public Function ParseLatex(ByVal text As String, ByVal isBlock As Boolean, ByRef outRoot As Long, ByRef outIsAmbiguous As Boolean, ByRef outEqTag As String) As Long
    ModTexAst.ResetArena isBlock
    ModTexLexer.TokenizeLatex text
    outRoot = ParseExpression(0)
    If ModTexAst.CurTok().TokType <> TOK_EOF Then
        ModTexAst.SetAmbiguous True
    End If
    outIsAmbiguous = ModTexAst.mIsAmbiguous
    outEqTag = ModTexAst.mEqTag
    ParseLatex = outRoot
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
