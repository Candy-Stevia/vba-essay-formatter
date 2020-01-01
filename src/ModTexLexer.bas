Attribute VB_Name = "ModTexLexer"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexLexer
' Purpose: Tokenizer for LaTeX mathematical expressions
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Const TOK_CMD As Integer = 1
Public Const TOK_LBRACE As Integer = 2
Public Const TOK_RBRACE As Integer = 3
Public Const TOK_LBRACK As Integer = 4
Public Const TOK_RBRACK As Integer = 5
Public Const TOK_LPAREN As Integer = 6
Public Const TOK_RPAREN As Integer = 7
Public Const TOK_SUB As Integer = 8
Public Const TOK_SUP As Integer = 9
Public Const TOK_AMP As Integer = 10
Public Const TOK_NEWLINE As Integer = 11
Public Const TOK_CHAR As Integer = 12
Public Const TOK_ESC_CHAR As Integer = 13
Public Const TOK_SPACE As Integer = 14
Public Const TOK_EOF As Integer = 15

Public Type TexToken
    TokType As Integer
    Val As String
    Pos As Long
End Type

Private Function IsAlpha(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsAlpha = False
        Exit Function
    End If
    c = AscW(ch) And &HFFFF&
    If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
        IsAlpha = True
    Else
        IsAlpha = False
    End If
End Function

Private Function IsSpaceChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then
        IsSpaceChar = False
        Exit Function
    End If
    If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = ChrW(11) Or ch = ChrW(&HA0) Then
        IsSpaceChar = True
    Else
        IsSpaceChar = False
    End If
End Function

Private Function IsTexDimension(ByVal s As String) As Boolean
    Dim n As Long
    Dim i As Long
    Dim ch As String
    Dim hasDigit As Boolean
    Dim hasDot As Boolean
    Dim unitStr As String

    n = Len(s)
    If n < 3 Then
        IsTexDimension = False
        Exit Function
    End If
    i = 1
    ch = Mid$(s, i, 1)
    If ch = "+" Or ch = "-" Then
        i = i + 1
    End If
    hasDigit = False
    hasDot = False
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            hasDigit = True
            i = i + 1
        ElseIf ch = "." And Not hasDot Then
            hasDot = True
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If Not hasDigit Then
        IsTexDimension = False
        Exit Function
    End If
    Do While i <= n
        If Mid$(s, i, 1) = " " Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    unitStr = LCase$(Mid$(s, i))
    If Len(unitStr) <> 2 Then
        IsTexDimension = False
    ElseIf InStr(1, "|pt|pc|in|bp|cm|mm|dd|cc|sp|em|ex|mu|", "|" & unitStr & "|") > 0 Then
        IsTexDimension = True
    Else
        IsTexDimension = False
    End If
End Function

Private Function SkipLineBreakOptArg(ByVal text As String, ByVal pos As Long, ByVal textLen As Long) As Long
    Dim p As Long
    Dim posAfterStar As Long
    Dim closePos As Long
    Dim inner As String

    p = pos
    If p <= textLen Then
        If Mid$(text, p, 1) = "*" Then
            p = p + 1
        End If
    End If
    posAfterStar = p
    Do While p <= textLen
        If IsSpaceChar(Mid$(text, p, 1)) Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop
    If p > textLen Then
        SkipLineBreakOptArg = posAfterStar
        Exit Function
    End If
    If Mid$(text, p, 1) <> "[" Then
        SkipLineBreakOptArg = posAfterStar
        Exit Function
    End If
    closePos = InStr(p + 1, text, "]")
    If closePos = 0 Or closePos - p > 24 Then
        SkipLineBreakOptArg = posAfterStar
        Exit Function
    End If
    inner = Trim$(Mid$(text, p + 1, closePos - p - 1))
    If IsTexDimension(inner) Then
        SkipLineBreakOptArg = closePos + 1
    Else
        SkipLineBreakOptArg = posAfterStar
    End If
End Function

Public Sub TokenizeLatex(ByVal text As String)
    Dim textLen As Long
    Dim pos As Long
    Dim startPos As Long
    Dim cmdStart As Long
    Dim ch As String
    Dim nextCh As String
    Dim cmdLetters As String
    Dim code As Long
    Dim nextCode As Long

    textLen = Len(text)
    pos = 1

    Do While pos <= textLen
        If ModTexAst.mTokenCount >= ModTexAst.MAX_TOKENS Then
            ModTexAst.SetAmbiguous True
            Exit Do
        End If
        ch = Mid$(text, pos, 1)

        If IsSpaceChar(ch) Then
            startPos = pos
            Do While pos <= textLen
                If IsSpaceChar(Mid$(text, pos, 1)) Then
                    pos = pos + 1
                Else
                    Exit Do
                End If
            Loop
            ModTexAst.AddToken TOK_SPACE, " ", startPos
        ElseIf ch = "%" Then
            pos = pos + 1
            Do While pos <= textLen
                ch = Mid$(text, pos, 1)
                If ch = vbCr Or ch = vbLf Or ch = ChrW(11) Then
                    pos = pos + 1
                    If ch = vbCr And pos <= textLen Then
                        If Mid$(text, pos, 1) = vbLf Then
                            pos = pos + 1
                        End If
                    End If
                    Exit Do
                End If
                pos = pos + 1
            Loop
        ElseIf ch = "\" Then
            startPos = pos
            pos = pos + 1
            If pos > textLen Then
                ModTexAst.AddToken TOK_CHAR, "\", startPos
            Else
                nextCh = Mid$(text, pos, 1)
                If nextCh = "\" Then
                    pos = pos + 1
                    ModTexAst.AddToken TOK_NEWLINE, "\\", startPos
                    pos = SkipLineBreakOptArg(text, pos, textLen)
                ElseIf InStr(1, "%&$#_{} ", nextCh) > 0 Then
                    pos = pos + 1
                    ModTexAst.AddToken TOK_ESC_CHAR, nextCh, startPos
                ElseIf InStr(1, ":,;!|", nextCh) > 0 Then
                    pos = pos + 1
                    ModTexAst.AddToken TOK_CMD, "\" & nextCh, startPos
                ElseIf IsAlpha(nextCh) Then
                    cmdStart = pos
                    pos = pos + 1
                    Do While pos <= textLen
                        If IsAlpha(Mid$(text, pos, 1)) Then
                            pos = pos + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    cmdLetters = Mid$(text, cmdStart, pos - cmdStart)
                    If (cmdLetters = "operatorname" Or cmdLetters = "tag") And pos <= textLen Then
                        If Mid$(text, pos, 1) = "*" Then
                            cmdLetters = cmdLetters & "*"
                            pos = pos + 1
                        End If
                    End If
                    ModTexAst.AddToken TOK_CMD, "\" & cmdLetters, startPos
                Else
                    code = AscW(nextCh) And &HFFFF&
                    If code >= &HD800& And code <= &HDBFF& And pos < textLen Then
                        nextCode = AscW(Mid$(text, pos + 1, 1)) And &HFFFF&
                        If nextCode >= &HDC00& And nextCode <= &HDFFF& Then
                            ModTexAst.AddToken TOK_CHAR, Mid$(text, pos, 2), startPos
                            pos = pos + 2
                        Else
                            ModTexAst.AddToken TOK_CHAR, nextCh, startPos
                            pos = pos + 1
                        End If
                    Else
                        ModTexAst.AddToken TOK_CHAR, nextCh, startPos
                        pos = pos + 1
                    End If
                End If
            End If
        ElseIf ch = "{" Then
            ModTexAst.AddToken TOK_LBRACE, "{", pos
            pos = pos + 1
        ElseIf ch = "}" Then
            ModTexAst.AddToken TOK_RBRACE, "}", pos
            pos = pos + 1
        ElseIf ch = "[" Then
            ModTexAst.AddToken TOK_LBRACK, "[", pos
            pos = pos + 1
        ElseIf ch = "]" Then
            ModTexAst.AddToken TOK_RBRACK, "]", pos
            pos = pos + 1
        ElseIf ch = "(" Then
            ModTexAst.AddToken TOK_LPAREN, "(", pos
            pos = pos + 1
        ElseIf ch = ")" Then
            ModTexAst.AddToken TOK_RPAREN, ")", pos
            pos = pos + 1
        ElseIf ch = "_" Then
            ModTexAst.AddToken TOK_SUB, "_", pos
            pos = pos + 1
        ElseIf ch = "^" Then
            ModTexAst.AddToken TOK_SUP, "^", pos
            pos = pos + 1
        ElseIf ch = "&" Then
            ModTexAst.AddToken TOK_AMP, "&", pos
            pos = pos + 1
        ElseIf ch = "$" Then
            ModTexAst.SetAmbiguous True
            pos = pos + 1
        Else
            code = AscW(ch) And &HFFFF&
            If code >= &HD800& And code <= &HDBFF& And pos < textLen Then
                nextCode = AscW(Mid$(text, pos + 1, 1)) And &HFFFF&
                If nextCode >= &HDC00& And nextCode <= &HDFFF& Then
                    ModTexAst.AddToken TOK_CHAR, Mid$(text, pos, 2), pos
                    pos = pos + 2
                Else
                    ModTexAst.AddToken TOK_CHAR, ch, pos
                    pos = pos + 1
                End If
            Else
                ModTexAst.AddToken TOK_CHAR, ch, pos
                pos = pos + 1
            End If
        End If
    Loop

    ModTexAst.AddToken TOK_EOF, "", pos
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
