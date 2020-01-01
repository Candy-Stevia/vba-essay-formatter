Attribute VB_Name = "ModTexAst"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexAst
' Purpose: AST node arena, parser state, token navigation, and group parser
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Const AST_CHAR As Integer = 1
Public Const AST_TEXT As Integer = 2
Public Const AST_FRAC As Integer = 3
Public Const AST_RAD As Integer = 4
Public Const AST_SUBSUP As Integer = 5
Public Const AST_NARY As Integer = 6
Public Const AST_DELIM As Integer = 7
Public Const AST_MATRIX As Integer = 8
Public Const AST_ACC As Integer = 9
Public Const AST_LIMLOW As Integer = 10
Public Const AST_LIMUPP As Integer = 11
Public Const AST_BORDERBOX As Integer = 12
Public Const AST_LIST As Integer = 13
Public Const AST_BAR As Integer = 14
Public Const AST_GROUPCHR As Integer = 15

Public Const MAX_TOKENS As Long = 10000
Public Const MAX_NODES As Long = 10000
Public Const MAX_DEPTH As Long = 64

Public Type AstNode
    NodeType As Integer
    Str1 As String
    Str2 As String
    Bool1 As Boolean
    Child1 As Long
    Child2 As Long
    Child3 As Long
    ChildList As Collection
End Type

Public mNodes() As AstNode
Public mNodeCount As Long
Public mTokens() As TexToken
Public mTokenCount As Long
Public mPos As Long
Public mDepth As Long
Public mInTextMode As Boolean
Public mIsBlock As Boolean
Public mIsAmbiguous As Boolean
Public mEqTag As String

Public Sub ResetArena(ByVal isBlock As Boolean)
    mNodeCount = 0
    ReDim mNodes(0 To 256)
    mTokenCount = 0
    ReDim mTokens(1 To 256)
    mPos = 1
    mDepth = 0
    mInTextMode = False
    mIsBlock = isBlock
    mIsAmbiguous = False
    mEqTag = ""
End Sub

Public Sub ReleaseArena()
    Erase mNodes
    Erase mTokens
    mNodeCount = 0
    mTokenCount = 0
    mPos = 0
    mDepth = 0
End Sub

Public Sub AddToken(ByVal tType As Integer, ByVal valStr As String, ByVal posVal As Long)
    Dim newCap As Long
    If mTokenCount >= MAX_TOKENS Then
        mIsAmbiguous = True
        If tType = TOK_EOF Then
            mTokens(MAX_TOKENS).TokType = TOK_EOF
            mTokens(MAX_TOKENS).Val = ""
            mTokens(MAX_TOKENS).Pos = posVal
        End If
        Exit Sub
    End If
    mTokenCount = mTokenCount + 1
    If mTokenCount > UBound(mTokens) Then
        newCap = UBound(mTokens) * 2
        If newCap > MAX_TOKENS Then newCap = MAX_TOKENS
        ReDim Preserve mTokens(1 To newCap)
    End If
    mTokens(mTokenCount).TokType = tType
    mTokens(mTokenCount).Val = valStr
    mTokens(mTokenCount).Pos = posVal
End Sub

Public Sub SetAmbiguous(ByVal flag As Boolean)
    mIsAmbiguous = flag
End Sub

Public Function AllocNode(ByVal nType As Integer) As Long
    Dim newCap As Long
    If mNodeCount >= MAX_NODES Then
        mIsAmbiguous = True
        Set mNodes(0).ChildList = Nothing
        AllocNode = 0
        Exit Function
    End If
    mNodeCount = mNodeCount + 1
    If mNodeCount > UBound(mNodes) Then
        newCap = UBound(mNodes) * 2
        If newCap > MAX_NODES Then newCap = MAX_NODES
        ReDim Preserve mNodes(0 To newCap)
    End If
    mNodes(mNodeCount).NodeType = nType
    mNodes(mNodeCount).Str1 = ""
    mNodes(mNodeCount).Str2 = ""
    mNodes(mNodeCount).Bool1 = False
    mNodes(mNodeCount).Child1 = 0
    mNodes(mNodeCount).Child2 = 0
    mNodes(mNodeCount).Child3 = 0
    Set mNodes(mNodeCount).ChildList = Nothing
    AllocNode = mNodeCount
End Function

Private Sub SkipSpaces()
    If mInTextMode Then Exit Sub
    Do While mPos <= mTokenCount
        If mTokens(mPos).TokType <> TOK_SPACE Then Exit Do
        mPos = mPos + 1
    Loop
End Sub

Public Function CurTok() As TexToken
    Dim eofTok As TexToken

    If mTokenCount = 0 Then
        eofTok.TokType = TOK_EOF
        CurTok = eofTok
        Exit Function
    End If
    SkipSpaces
    If mPos <= mTokenCount Then
        CurTok = mTokens(mPos)
    Else
        CurTok = mTokens(mTokenCount)
    End If
End Function

Public Sub AdvanceTok()
    mPos = mPos + 1
    SkipSpaces
End Sub

Public Function ConsumeTok() As TexToken
    ConsumeTok = CurTok()
    AdvanceTok
End Function

Public Function ExpectTok(ByVal tType As Integer, Optional ByVal valStr As String = "") As TexToken
    Dim tok As TexToken
    tok = CurTok()
    If tok.TokType <> tType Or (Len(valStr) > 0 And tok.Val <> valStr) Then
        mIsAmbiguous = True
    End If
    AdvanceTok
    ExpectTok = tok
End Function

Public Function FindMatching(ByVal openType As Integer, ByVal closeType As Integer) As Long
    Dim depth As Long
    Dim i As Long
    depth = 0
    For i = mPos To mTokenCount
        If mTokens(i).TokType = openType Then
            depth = depth + 1
        ElseIf mTokens(i).TokType = closeType Then
            depth = depth - 1
            If depth = 0 Then
                FindMatching = i
                Exit Function
            End If
        End If
    Next i
    FindMatching = 0
End Function

Private Function SkipOneNonSpace(ByVal idx As Long) As Long
    Do While idx + 1 <= mTokenCount
        idx = idx + 1
        If mTokens(idx).TokType <> TOK_SPACE Then
            Exit Do
        End If
    Loop
    SkipOneNonSpace = idx
End Function

Public Function FindMatchingBar() As Long
    Dim i As Long, dParen As Long, dBrack As Long, dBrace As Long, dLeft As Long, dEnv As Long
    Dim t As TexToken
    Dim atTop As Boolean

    i = mPos + 1
    Do While i <= mTokenCount
        t = mTokens(i)
        atTop = (dParen = 0 And dBrack = 0 And dBrace = 0 And dLeft = 0 And dEnv = 0)
        Select Case t.TokType
            Case TOK_EOF
                Exit Function
            Case TOK_LPAREN
                dParen = dParen + 1
            Case TOK_RPAREN
                If dParen = 0 Then
                    Exit Function
                End If
                dParen = dParen - 1
            Case TOK_LBRACK
                dBrack = dBrack + 1
            Case TOK_RBRACK
                If dBrack = 0 Then
                    Exit Function
                End If
                dBrack = dBrack - 1
            Case TOK_LBRACE
                dBrace = dBrace + 1
            Case TOK_RBRACE
                If dBrace = 0 Then
                    Exit Function
                End If
                dBrace = dBrace - 1
            Case TOK_AMP, TOK_NEWLINE
                If dEnv = 0 Then
                    Exit Function
                End If
            Case TOK_CMD
                If t.Val = "\left" Then
                    dLeft = dLeft + 1
                    i = SkipOneNonSpace(i)
                ElseIf t.Val = "\right" Then
                    If dLeft = 0 Then
                        Exit Function
                    End If
                    dLeft = dLeft - 1
                    i = SkipOneNonSpace(i)
                ElseIf t.Val = "\begin" Then
                    dEnv = dEnv + 1
                ElseIf t.Val = "\end" Then
                    If dEnv = 0 Then
                        Exit Function
                    End If
                    dEnv = dEnv - 1
                ElseIf atTop Then
                    If InStr(1, "|\le|\ge|\leq|\geq|\neq|\approx|\in|", "|" & t.Val & "|") > 0 Then
                        Exit Function
                    End If
                End If
            Case TOK_CHAR
                If atTop Then
                    If InStr(1, "<>=,", t.Val) > 0 Then
                        Exit Function
                    ElseIf t.Val = "|" Then
                        FindMatchingBar = i
                        Exit Function
                    End If
                End If
        End Select
        i = i + 1
    Loop
End Function

Public Function ConsumeLiteralTextArg() As String
    Dim tok As TexToken
    Dim depth As Long
    Dim res As String, cmd As String, sym As String

    If CurTok().TokType <> TOK_LBRACE Then
        tok = ConsumeTok()
        ConsumeLiteralTextArg = tok.Val
        Exit Function
    End If
    mInTextMode = True
    AdvanceTok
    depth = 1
    res = ""
    Do While CurTok().TokType <> TOK_EOF
        If CurTok().TokType = TOK_LBRACE Then
            depth = depth + 1
            res = res & "{"
            AdvanceTok
        ElseIf CurTok().TokType = TOK_RBRACE Then
            depth = depth - 1
            If depth = 0 Then
                mInTextMode = False
                AdvanceTok
                Exit Do
            End If
            AdvanceTok
            res = res & "}"
        ElseIf CurTok().TokType = TOK_SPACE Then
            AdvanceTok
            res = res & " "
        ElseIf CurTok().TokType = TOK_ESC_CHAR Then
            tok = ConsumeTok()
            res = res & tok.Val
        ElseIf CurTok().TokType = TOK_CMD Then
            tok = ConsumeTok()
            cmd = tok.Val
            If cmd = "\ " Or cmd = "\," Or cmd = "\:" Or cmd = "\;" Or cmd = "\quad" Or cmd = "\qquad" Then
                res = res & " "
            ElseIf cmd = "\%" Or cmd = "\$" Or cmd = "\&" Or cmd = "\_" Or cmd = "\#" Or cmd = "\{" Or cmd = "\}" Then
                res = res & Mid$(cmd, 2)
            ElseIf cmd = "\u" Or cmd = "\v" Or cmd = "\t" Or cmd = "\r" Or cmd = "\b" Or cmd = "\c" Or cmd = "\d" Or cmd = "\k" Or cmd = "\H" Then
                ' Drop LaTeX accent command to avoid raw leak
            Else
                sym = ModTexSymbols.LookupGreek(cmd)
                If Len(sym) = 0 Then
                    sym = ModTexSymbols.LookupSymbol(cmd)
                End If
                If Len(sym) > 0 Then
                    res = res & sym
                Else
                    res = res & Mid$(cmd, 2)
                End If
            End If
        Else
            tok = ConsumeTok()
            res = res & tok.Val
        End If
    Loop
    mInTextMode = False
    If depth > 0 Then mIsAmbiguous = True
    ConsumeLiteralTextArg = res
End Function

Public Sub SkipLimitMods()
    Do While CurTok().TokType = TOK_CMD
        If CurTok().Val <> "\limits" And CurTok().Val <> "\nolimits" Then Exit Do
        AdvanceTok
    Loop
End Sub

Public Function HandleSubSup(ByVal baseNode As Long) As Long
    Dim subNode As Long, supNode As Long, resNode As Long, extraSup As Long, listNd As Long, isPrimeSup As Boolean, primes As String
    Dim supItems As Collection
    Do While CurTok().TokType = TOK_SUB Or CurTok().TokType = TOK_SUP Or (CurTok().TokType = TOK_CHAR And CurTok().Val = "'")
        If CurTok().TokType = TOK_SUB Then
            If subNode <> 0 Then Exit Do
            AdvanceTok
            subNode = ModTexParser.ParseArgument()
        ElseIf CurTok().TokType = TOK_SUP Then
            If supNode <> 0 And Not isPrimeSup Then Exit Do
            AdvanceTok
            extraSup = ModTexParser.ParseArgument()
            If isPrimeSup Then
                Set supItems = New Collection
                supItems.Add supNode
                If extraSup <> 0 Then supItems.Add extraSup
                listNd = AllocNode(AST_LIST)
                Set mNodes(listNd).ChildList = supItems
                supNode = listNd
                isPrimeSup = False
            Else
                supNode = extraSup
            End If
        Else
            If supNode <> 0 Then Exit Do
            primes = ""
            Do While CurTok().TokType = TOK_CHAR And CurTok().Val = "'"
                primes = primes & ChrW(&H2032)
                AdvanceTok
            Loop
            supNode = AllocNode(AST_CHAR)
            mNodes(supNode).Str1 = primes
            isPrimeSup = True
        End If
    Loop
    If subNode <> 0 Or supNode <> 0 Then
        resNode = AllocNode(AST_SUBSUP)
        mNodes(resNode).Child1 = baseNode
        mNodes(resNode).Child2 = subNode
        mNodes(resNode).Child3 = supNode
        HandleSubSup = resNode
    Else
        HandleSubSup = baseNode
    End If
End Function

Public Function ParseGroup() As Long
    ParseGroup = ModTexParser.ParseGroupCore()
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
