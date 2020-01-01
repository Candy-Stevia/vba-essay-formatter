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
Public mInTextMode As Boolean
Public mIsBlock As Boolean
Public mIsAmbiguous As Boolean
Public mEqTag As String

Public Sub ResetArena(ByVal isBlock As Boolean)
    mNodeCount = 0
    ReDim mNodes(1 To 256)
    mTokenCount = 0
    ReDim mTokens(1 To 256)
    mPos = 1
    mInTextMode = False
    mIsBlock = isBlock
    mIsAmbiguous = False
    mEqTag = ""
End Sub

Public Sub AddToken(ByVal tType As Integer, ByVal valStr As String, ByVal posVal As Long)
    mTokenCount = mTokenCount + 1
    If mTokenCount > UBound(mTokens) Then
        ReDim Preserve mTokens(1 To UBound(mTokens) + 128)
    End If
    mTokens(mTokenCount).TokType = tType
    mTokens(mTokenCount).Val = valStr
    mTokens(mTokenCount).Pos = posVal
End Sub

Public Sub SetAmbiguous(ByVal flag As Boolean)
    mIsAmbiguous = flag
End Sub

Public Function AllocNode(ByVal nType As Integer) As Long
    mNodeCount = mNodeCount + 1
    If mNodeCount > UBound(mNodes) Then
        ReDim Preserve mNodes(1 To UBound(mNodes) + 128)
    End If
    mNodes(mNodeCount).NodeType = nType
    mNodes(mNodeCount).Str1 = ""
    mNodes(mNodeCount).Str2 = ""
    mNodes(mNodeCount).Bool1 = False
    mNodes(mNodeCount).Child1 = 0
    mNodes(mNodeCount).Child2 = 0
    mNodes(mNodeCount).Child3 = 0
    Set mNodes(mNodeCount).ChildList = New Collection
    AllocNode = mNodeCount
End Function

Public Function CurTok() As TexToken
    If Not mInTextMode Then
        Do While mPos <= mTokenCount
            If mTokens(mPos).TokType = TOK_SPACE Then
                mPos = mPos + 1
            Else
                Exit Do
            End If
        Loop
    End If
    If mPos <= mTokenCount Then
        CurTok = mTokens(mPos)
    Else
        CurTok = mTokens(mTokenCount)
    End If
End Function

Public Sub AdvanceTok()
    mPos = mPos + 1
    If Not mInTextMode Then
        Do While mPos <= mTokenCount
            If mTokens(mPos).TokType = TOK_SPACE Then
                mPos = mPos + 1
            Else
                Exit Do
            End If
        Loop
    End If
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

Public Function FindMatchingBar() As Long
    Dim i As Long
    Dim t As TexToken
    For i = mPos + 1 To mTokenCount
        t = mTokens(i)
        If t.TokType = TOK_CHAR Then
            If InStr(1, "<>=,", t.Val) > 0 Then
                FindMatchingBar = 0
                Exit Function
            ElseIf t.Val = "|" Then
                FindMatchingBar = i
                Exit Function
            End If
        ElseIf t.TokType = TOK_CMD Then
            If t.Val = "\le" Or t.Val = "\ge" Or t.Val = "\leq" Or t.Val = "\geq" Or t.Val = "\neq" Or t.Val = "\approx" Or t.Val = "\in" Then
                FindMatchingBar = 0
                Exit Function
            End If
        ElseIf t.TokType = TOK_AMP Or t.TokType = TOK_NEWLINE Or t.TokType = TOK_EOF Then
            FindMatchingBar = 0
            Exit Function
        End If
    Next i
    FindMatchingBar = 0
End Function

Public Function ConsumeLiteralTextArg() As String
    Dim tok As TexToken
    Dim depth As Long
    Dim res As String
    Dim cmd As String

    If CurTok().TokType <> TOK_LBRACE Then
        tok = ConsumeTok()
        ConsumeLiteralTextArg = tok.Val
        Exit Function
    End If

    AdvanceTok
    mInTextMode = True
    depth = 1
    res = ""

    Do While CurTok().TokType <> TOK_EOF
        If CurTok().TokType = TOK_LBRACE Then
            depth = depth + 1
            res = res & "{"
            AdvanceTok
        ElseIf CurTok().TokType = TOK_RBRACE Then
            depth = depth - 1
            AdvanceTok
            If depth = 0 Then
                Exit Do
            End If
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
                res = res & Mid$(cmd, 2)
            End If
        Else
            tok = ConsumeTok()
            res = res & tok.Val
        End If
    Loop

    mInTextMode = False
    ConsumeLiteralTextArg = res
End Function

Public Function HandleSubSup(ByVal baseNode As Long) As Long
    Dim subNode As Long
    Dim supNode As Long
    Dim resNode As Long

    subNode = 0
    supNode = 0

    Do While CurTok().TokType = TOK_SUB Or CurTok().TokType = TOK_SUP
        If CurTok().TokType = TOK_SUB Then
            If subNode <> 0 Then
                Exit Do
            End If
            AdvanceTok
            subNode = ModTexParser.ParseArgument()
        ElseIf CurTok().TokType = TOK_SUP Then
            If supNode <> 0 Then
                Exit Do
            End If
            AdvanceTok
            supNode = ModTexParser.ParseArgument()
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
    Dim items As Collection
    Dim pNode As Long
    Dim listNode As Long

    ExpectTok TOK_LBRACE
    Set items = New Collection

    Do While CurTok().TokType <> TOK_EOF And CurTok().TokType <> TOK_RBRACE
        pNode = ModTexParser.ParsePrimary()
        If pNode <> 0 Then
            items.Add pNode
        Else
            If CurTok().TokType <> TOK_EOF And CurTok().TokType <> TOK_RBRACE Then
                AdvanceTok
            End If
        End If
    Loop

    ExpectTok TOK_RBRACE
    listNode = AllocNode(AST_LIST)
    Set mNodes(listNode).ChildList = items
    ParseGroup = listNode
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
