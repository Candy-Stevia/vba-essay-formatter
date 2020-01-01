Attribute VB_Name = "ModTexSymbols"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexSymbols
' Purpose: Symbol lookup tables for Greek letters, operators, accents, and functions
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Private mGreekDict As Object
Private mSymDict As Object

Private Sub AddDictEntries(ByVal dict As Object, ByVal spec As String)
    Dim items() As String
    Dim i As Long
    Dim sepPos As Long
    Dim k As String
    Dim h As String

    items = Split(spec, ",")
    For i = LBound(items) To UBound(items)
        sepPos = InStr(1, items(i), ":")
        If sepPos > 1 Then
            k = "\" & Left$(items(i), sepPos - 1)
            h = Mid$(items(i), sepPos + 1)
            dict(k) = ChrW(CLng("&H" & h))
        End If
    Next i
End Sub

Public Function LookupGreek(ByVal cmd As String) As String
    Dim s1 As String
    Dim s2 As String

    If mGreekDict Is Nothing Then
        Set mGreekDict = CreateObject("Scripting.Dictionary")
        s1 = "alpha:3B1,beta:3B2,gamma:3B3,delta:3B4,epsilon:3F5,varepsilon:3B5,zeta:3B6,eta:3B7,theta:3B8,vartheta:3D1,iota:3B9,kappa:3BA,lambda:3BB,mu:3BC,nu:3BD,xi:3BE,pi:3C0,varpi:3D6,rho:3C1,varrho:3F1"
        s2 = "sigma:3C3,varsigma:3C2,tau:3C4,upsilon:3C5,phi:3D5,varphi:3C6,chi:3C7,psi:3C8,omega:3C9,Gamma:393,Delta:394,Theta:398,Lambda:39B,Xi:39E,Pi:3A0,Sigma:3A3,Upsilon:3A5,Phi:3A6,Psi:3A8,Omega:3A9"
        AddDictEntries mGreekDict, s1
        AddDictEntries mGreekDict, s2
    End If
    If mGreekDict.Exists(cmd) Then
        LookupGreek = mGreekDict(cmd)
    Else
        LookupGreek = ""
    End If
End Function

Public Function LookupSymbol(ByVal cmd As String) As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String

    If mSymDict Is Nothing Then
        Set mSymDict = CreateObject("Scripting.Dictionary")
        s1 = "leq:2264,le:2264,geq:2265,ge:2265,neq:2260,ne:2260,approx:2248,times:D7,div:F7,cdot:22C5,pm:B1,mp:2213,in:2208,notin:2209,to:2192,rightarrow:2192,leftarrow:2190,gets:2190,Rightarrow:21D2,Leftarrow:21D0"
        s2 = "Leftrightarrow:21D4,infty:221E,partial:2202,nabla:2207,ldots:2026,dots:2026,cdots:22EF,vdots:22EE,ddots:22F1,mid:2223,prime:2032,sim:223C,forall:2200,exists:2203,nexists:2204,subset:2282,subseteq:2286,supset:2283,supseteq:2287"
        s3 = "cup:222A,cap:2229,langle:27E8,rangle:27E9,lfloor:230A,rfloor:230B,lceil:2308,rceil:2309,circ:2218,oplus:2295,otimes:2297,perp:22A5,equiv:2261,propto:221D,cong:2245,simeq:2243,triangle:25B3,emptyset:2205,varnothing:2205,setminus:2216"
        s4 = "vert:7C,leftrightarrow:2194,mapsto:21A6,implies:27F9,iff:27FA,land:2227,wedge:2227,lor:2228,vee:2228,neg:AC,lnot:AC,ell:2113,hbar:210F,ll:226A,gg:226B"
        AddDictEntries mSymDict, s1
        AddDictEntries mSymDict, s2
        AddDictEntries mSymDict, s3
        AddDictEntries mSymDict, s4
    End If
    If mSymDict.Exists(cmd) Then
        LookupSymbol = mSymDict(cmd)
    Else
        LookupSymbol = ""
    End If
End Function

Public Function LookupNotNegation(ByVal tokVal As String) As String
    Select Case tokVal
        Case "\in"
            LookupNotNegation = ChrW(&H2209)
        Case "="
            LookupNotNegation = ChrW(&H2260)
        Case "\equiv"
            LookupNotNegation = ChrW(&H2262)
        Case "\subset"
            LookupNotNegation = ChrW(&H2284)
        Case "\subseteq"
            LookupNotNegation = ChrW(&H2288)
        Case "\approx"
            LookupNotNegation = ChrW(&H2249)
        Case "\sim"
            LookupNotNegation = ChrW(&H2241)
        Case "<"
            LookupNotNegation = ChrW(&H226E)
        Case ">"
            LookupNotNegation = ChrW(&H226F)
        Case "\leq", "\le"
            LookupNotNegation = ChrW(&H2270)
        Case "\geq", "\ge"
            LookupNotNegation = ChrW(&H2271)
        Case Else
            LookupNotNegation = ""
    End Select
End Function

Public Function LookupNary(ByVal cmd As String) As String
    Select Case cmd
        Case "\sum"
            LookupNary = ChrW(&H2211)
        Case "\int"
            LookupNary = ChrW(&H222B)
        Case "\iint"
            LookupNary = ChrW(&H222C)
        Case "\iiint"
            LookupNary = ChrW(&H222D)
        Case "\oint"
            LookupNary = ChrW(&H222E)
        Case "\prod"
            LookupNary = ChrW(&H220F)
        Case "\coprod"
            LookupNary = ChrW(&H2210)
        Case "\bigcup"
            LookupNary = ChrW(&H22C3)
        Case "\bigcap"
            LookupNary = ChrW(&H22C2)
        Case Else
            LookupNary = ""
    End Select
End Function

Public Function LookupAccent(ByVal cmd As String) As String
    Select Case cmd
        Case "\hat", "\widehat"
            LookupAccent = ChrW(&H302)
        Case "\bar"
            LookupAccent = ChrW(&H304)
        Case "\vec"
            LookupAccent = ChrW(&H20D7)
        Case "\tilde", "\widetilde"
            LookupAccent = ChrW(&H303)
        Case "\check"
            LookupAccent = ChrW(&H30C)
        Case "\dot"
            LookupAccent = ChrW(&H307)
        Case "\ddot"
            LookupAccent = ChrW(&H308)
        Case Else
            LookupAccent = ""
    End Select
End Function

Public Function IsNamedFunction(ByVal nameStr As String) As Boolean
    Select Case nameStr
        Case "sin", "cos", "tan", "cot", "sec", "csc"
            IsNamedFunction = True
        Case "arcsin", "arccos", "arctan"
            IsNamedFunction = True
        Case "sinh", "cosh", "tanh"
            IsNamedFunction = True
        Case "ln", "log", "exp"
            IsNamedFunction = True
        Case "lim", "min", "max", "inf", "sup"
            IsNamedFunction = True
        Case "deg", "dim", "det", "ker", "hom"
            IsNamedFunction = True
        Case "arg", "argmin", "argmax"
            IsNamedFunction = True
        Case Else
            IsNamedFunction = False
    End Select
End Function

Public Function IsOperatorChar(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    code = AscW(ch) And &HFFFF&
    If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
        Exit Function
    End If
    Select Case code
        Case 42, 43, 44, 45, 47, 60, 61, 62 ' * + , - / < = >
            IsOperatorChar = True
        Case &H2212&, &HD7&, &HF7&, &H22C5&, &HB1&, &H2213& ' minus, times, div, cdot, pm, mp
            IsOperatorChar = True
        Case &H2264&, &H2265&, &H2260&, &H2248& ' le, ge, ne, approx
            IsOperatorChar = True
        Case &H2208&, &H2209&, &H2192& ' in, notin, rightarrow
            IsOperatorChar = True
        Case &H2223&, &H2225&, &H223C& ' mid, parallel, sim
            IsOperatorChar = True
        Case &H2218&, &H2295&, &H2297&, &H22A5& ' circ, oplus, otimes, perp
            IsOperatorChar = True
        Case &H2261&, &H221D&, &H2245&, &H2243&, &H2216& ' equiv, propto, cong, simeq, setminus
            IsOperatorChar = True
        Case &H222A&, &H2229&, &H2282&, &H2286&, &H2283&, &H2287& ' cup, cap, sub, sube, sup, supe
            IsOperatorChar = True
        Case &H2262&, &H2284&, &H2288&, &H2249&, &H2241& ' nequiv, nsub, nsube, napprox, nsim
            IsOperatorChar = True
        Case &H226E&, &H226F&, &H2270&, &H2271& ' nless, ngtr, nleq, ngeq
            IsOperatorChar = True
        Case &H2190&, &H2194&, &H21A6&, &H21D0&, &H21D2&, &H21D4&, &H27F9&, &H27FA& ' arrows
            IsOperatorChar = True
        Case &H2227&, &H2228&, &HAC&, &H226A&, &H226B& ' land, lor, neg, ll, gg
            IsOperatorChar = True
        Case Else
            IsOperatorChar = False
    End Select
End Function

Public Function LookupSpace(ByVal cmd As String) As String
    Select Case cmd
        Case "\,"
            LookupSpace = ChrW(&H2009)
        Case "\;", "\:", "\ "
            LookupSpace = ChrW(&H2005)
        Case "\quad"
            LookupSpace = ChrW(&H2001)
        Case "\qquad"
            LookupSpace = ChrW(&H2001) & ChrW(&H2001)
        Case Else
            LookupSpace = vbNullString
    End Select
End Function

Public Function LookupFontMode(ByVal cmd As String) As String
    Select Case cmd
        Case "\text", "\mbox", "\textrm"
            LookupFontMode = "text"
        Case "\mathrm", "\operatorname"
            LookupFontMode = "roman"
        Case "\mathbb"
            LookupFontMode = "double-struck"
        Case "\mathbf", "\bm", "\boldsymbol"
            LookupFontMode = "bold"
        Case "\mathcal"
            LookupFontMode = "script"
        Case Else
            LookupFontMode = vbNullString
    End Select
End Function

Public Function EscapeXml(ByVal s As String) As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    Dim code As Long
    Dim nextCode As Long
    Dim parts() As String
    Dim pCount As Long
    Dim needsWork As Boolean

    n = Len(s)
    If n = 0 Then
        EscapeXml = ""
        Exit Function
    End If
    If n = 1 Then
        code = AscW(s) And &HFFFF&
        Select Case code
            Case 38
                EscapeXml = "&amp;"
            Case 60
                EscapeXml = "&lt;"
            Case 62
                EscapeXml = "&gt;"
            Case 34
                EscapeXml = "&quot;"
            Case 39
                EscapeXml = "&apos;"
            Case 9, 10, 13
                EscapeXml = s
            Case Is < 32, &HFFFE&, &HFFFF&, &HD800& To &HDFFF&
                EscapeXml = ""
            Case Else
                EscapeXml = s
        End Select
        Exit Function
    End If
    needsWork = False
    i = 1
    Do While i <= n
        code = AscW(Mid$(s, i, 1)) And &HFFFF&
        If code = 38 Or code = 60 Or code = 62 Or code = 34 Or code = 39 Then
            needsWork = True
            Exit Do
        ElseIf code < 32 Then
            If code <> 9 And code <> 10 And code <> 13 Then
                needsWork = True
                Exit Do
            End If
        ElseIf code >= &HD800& And code <= &HDFFF& Then
            needsWork = True
            Exit Do
        ElseIf code = &HFFFE& Or code = &HFFFF& Then
            needsWork = True
            Exit Do
        End If
        i = i + 1
    Loop
    If Not needsWork Then
        EscapeXml = s
        Exit Function
    End If
    ReDim parts(1 To n)
    pCount = 0
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        If code = 38 Then
            pCount = pCount + 1
            parts(pCount) = "&amp;"
        ElseIf code = 60 Then
            pCount = pCount + 1
            parts(pCount) = "&lt;"
        ElseIf code = 62 Then
            pCount = pCount + 1
            parts(pCount) = "&gt;"
        ElseIf code = 34 Then
            pCount = pCount + 1
            parts(pCount) = "&quot;"
        ElseIf code = 39 Then
            pCount = pCount + 1
            parts(pCount) = "&apos;"
        ElseIf code = 9 Or code = 10 Or code = 13 Then
            pCount = pCount + 1
            parts(pCount) = ch
        ElseIf code < 32 Or code = &HFFFE& Or code = &HFFFF& Then
            ' Strip illegal XML 1.0 character
        ElseIf code >= &HD800& And code <= &HDBFF& Then
            If i < n Then
                nextCode = AscW(Mid$(s, i + 1, 1)) And &HFFFF&
                If nextCode >= &HDC00& And nextCode <= &HDFFF& Then
                    pCount = pCount + 1
                    parts(pCount) = Mid$(s, i, 2)
                    i = i + 1
                End If
            End If
        ElseIf code >= &HDC00& And code <= &HDFFF& Then
            ' Strip lone low surrogate
        Else
            pCount = pCount + 1
            parts(pCount) = ch
        End If
        i = i + 1
    Loop
    If pCount = 0 Then
        EscapeXml = ""
    ElseIf pCount = 1 Then
        EscapeXml = parts(1)
    Else
        ReDim Preserve parts(1 To pCount)
        EscapeXml = Join(parts, "")
    End If
End Function

Public Function IsNoOpMathCmd(ByVal cmd As String) As Boolean
    IsNoOpMathCmd = (cmd = "\limits" Or cmd = "\nolimits" Or cmd = "\nonumber" Or cmd = "\notag")
End Function

Public Function IsStyleMathCmd(ByVal cmd As String) As Boolean
    IsStyleMathCmd = (cmd = "\displaystyle" Or cmd = "\textstyle")
End Function

Public Function IsExpressionStopTok(ByVal tokType As Integer, ByVal tokVal As String) As Boolean
    If tokType = TOK_EOF Or tokType = TOK_RBRACE Or tokType = TOK_RPAREN Or tokType = TOK_RBRACK Or tokType = TOK_AMP Or tokType = TOK_NEWLINE Then
        IsExpressionStopTok = True
    ElseIf tokType = TOK_CMD Then
        IsExpressionStopTok = (tokVal = "\end" Or tokVal = "\right")
    End If
End Function

Public Function MapBigCloseCmd(ByVal cmd As String) As String
    Select Case cmd
        Case "\big", "\bigg", "\Big", "\Bigg"
            MapBigCloseCmd = cmd
        Case "\bigl"
            MapBigCloseCmd = "\bigr"
        Case "\Bigl"
            MapBigCloseCmd = "\Bigr"
        Case "\biggl"
            MapBigCloseCmd = "\biggr"
        Case "\Biggl"
            MapBigCloseCmd = "\Biggr"
    End Select
End Function

Public Function IsBigCloseOnlyCmd(ByVal cmd As String) As Boolean
    IsBigCloseOnlyCmd = (cmd = "\bigr" Or cmd = "\Bigr" Or cmd = "\biggr" Or cmd = "\Biggr")
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
