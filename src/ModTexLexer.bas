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
    Dim c As Integer
    If Len(ch) = 0 Then
        IsAlpha = False
        Exit Function
    End If
    c = Asc(ch)
    If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
        IsAlpha = True
    Else
        IsAlpha = False
    End If
End Function

Private Function IsSpaceChar(ByVal ch As String) As Boolean
    If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
        IsSpaceChar = True
    Else
        IsSpaceChar = False
    End If
End Function

Public Sub TokenizeLatex(ByVal text As String)
    Dim textLen As Long
    Dim pos As Long
    Dim startPos As Long
    Dim ch As String
    Dim nextCh As String
    Dim cmdLetters As String

    textLen = Len(text)
    pos = 1

    Do While pos <= textLen
        ch = Mid$(text, pos, 1)

        If IsSpaceChar(ch) Then
            startPos = pos
            Do While pos <= textLen And IsSpaceChar(Mid$(text, pos, 1))
                pos = pos + 1
            Loop
            ModTexAst.AddToken TOK_SPACE, " ", startPos
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
                ElseIf InStr(1, "%&$#_{} ", nextCh) > 0 Then
                    pos = pos + 1
                    ModTexAst.AddToken TOK_ESC_CHAR, nextCh, startPos
                ElseIf InStr(1, ":,;!|", nextCh) > 0 Then
                    pos = pos + 1
                    ModTexAst.AddToken TOK_CMD, "\" & nextCh, startPos
                ElseIf IsAlpha(nextCh) Then
                    cmdLetters = ""
                    Do While pos <= textLen And IsAlpha(Mid$(text, pos, 1))
                        cmdLetters = cmdLetters & Mid$(text, pos, 1)
                        pos = pos + 1
                    Loop
                    If pos <= textLen And Mid$(text, pos, 1) = "*" Then
                        cmdLetters = cmdLetters & "*"
                        pos = pos + 1
                    End If
                    ModTexAst.AddToken TOK_CMD, "\" & cmdLetters, startPos
                Else
                    ModTexAst.AddToken TOK_CHAR, nextCh, startPos
                    pos = pos + 1
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
        Else
            ModTexAst.AddToken TOK_CHAR, ch, pos
            pos = pos + 1
        End If
    Loop

    ModTexAst.AddToken TOK_EOF, "", pos
End Sub
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
