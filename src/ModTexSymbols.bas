Attribute VB_Name = "ModTexSymbols"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModTexSymbols
' Purpose: Symbol lookup tables for Greek letters, operators, accents, and functions
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Function LookupGreek(ByVal cmd As String) As String
    Select Case cmd
        Case "\alpha"
            LookupGreek = ChrW(&H3B1)
        Case "\beta"
            LookupGreek = ChrW(&H3B2)
        Case "\gamma"
            LookupGreek = ChrW(&H3B3)
        Case "\delta"
            LookupGreek = ChrW(&H3B4)
        Case "\epsilon"
            LookupGreek = ChrW(&H3F5)
        Case "\varepsilon"
            LookupGreek = ChrW(&H3B5)
        Case "\zeta"
            LookupGreek = ChrW(&H3B6)
        Case "\eta"
            LookupGreek = ChrW(&H3B7)
        Case "\theta"
            LookupGreek = ChrW(&H3B8)
        Case "\vartheta"
            LookupGreek = ChrW(&H3D1)
        Case "\iota"
            LookupGreek = ChrW(&H3B9)
        Case "\kappa"
            LookupGreek = ChrW(&H3BA)
        Case "\lambda"
            LookupGreek = ChrW(&H3BB)
        Case "\mu"
            LookupGreek = ChrW(&H3BC)
        Case "\nu"
            LookupGreek = ChrW(&H3BD)
        Case "\xi"
            LookupGreek = ChrW(&H3BE)
        Case "\pi"
            LookupGreek = ChrW(&H3C0)
        Case "\varpi"
            LookupGreek = ChrW(&H3D6)
        Case "\rho"
            LookupGreek = ChrW(&H3C1)
        Case "\varrho"
            LookupGreek = ChrW(&H3F1)
        Case "\sigma"
            LookupGreek = ChrW(&H3C3)
        Case "\varsigma"
            LookupGreek = ChrW(&H3C2)
        Case "\tau"
            LookupGreek = ChrW(&H3C4)
        Case "\upsilon"
            LookupGreek = ChrW(&H3C5)
        Case "\phi"
            LookupGreek = ChrW(&H3D5)
        Case "\varphi"
            LookupGreek = ChrW(&H3C6)
        Case "\chi"
            LookupGreek = ChrW(&H3C7)
        Case "\psi"
            LookupGreek = ChrW(&H3C8)
        Case "\omega"
            LookupGreek = ChrW(&H3C9)
        Case "\Gamma"
            LookupGreek = ChrW(&H393)
        Case "\Delta"
            LookupGreek = ChrW(&H394)
        Case "\Theta"
            LookupGreek = ChrW(&H398)
        Case "\Lambda"
            LookupGreek = ChrW(&H39B)
        Case "\Xi"
            LookupGreek = ChrW(&H39E)
        Case "\Pi"
            LookupGreek = ChrW(&H3A0)
        Case "\Sigma"
            LookupGreek = ChrW(&H3A3)
        Case "\Upsilon"
            LookupGreek = ChrW(&H3A5)
        Case "\Phi"
            LookupGreek = ChrW(&H3A6)
        Case "\Psi"
            LookupGreek = ChrW(&H3A8)
        Case "\Omega"
            LookupGreek = ChrW(&H3A9)
        Case Else
            LookupGreek = ""
    End Select
End Function

Public Function LookupSymbol(ByVal cmd As String) As String
    Select Case cmd
        Case "\leq", "\le"
            LookupSymbol = ChrW(&H2264)
        Case "\geq", "\ge"
            LookupSymbol = ChrW(&H2265)
        Case "\neq", "\ne"
            LookupSymbol = ChrW(&H2260)
        Case "\approx"
            LookupSymbol = ChrW(&H2248)
        Case "\times"
            LookupSymbol = ChrW(&HD7)
        Case "\div"
            LookupSymbol = ChrW(&HF7)
        Case "\cdot"
            LookupSymbol = ChrW(&H22C5)
        Case "\pm"
            LookupSymbol = ChrW(&HB1)
        Case "\mp"
            LookupSymbol = ChrW(&H2213)
        Case "\in"
            LookupSymbol = ChrW(&H2208)
        Case "\notin"
            LookupSymbol = ChrW(&H2209)
        Case "\to", "\rightarrow"
            LookupSymbol = ChrW(&H2192)
        Case "\leftarrow", "\gets"
            LookupSymbol = ChrW(&H2190)
        Case "\Rightarrow"
            LookupSymbol = ChrW(&H21D2)
        Case "\Leftarrow"
            LookupSymbol = ChrW(&H21D0)
        Case "\Leftrightarrow"
            LookupSymbol = ChrW(&H21D4)
        Case "\infty"
            LookupSymbol = ChrW(&H221E)
        Case "\partial"
            LookupSymbol = ChrW(&H2202)
        Case "\nabla"
            LookupSymbol = ChrW(&H2207)
        Case "\ldots", "\dots"
            LookupSymbol = ChrW(&H2026)
        Case "\cdots"
            LookupSymbol = ChrW(&H22EF)
        Case "\vdots"
            LookupSymbol = ChrW(&H22EE)
        Case "\ddots"
            LookupSymbol = ChrW(&H22F1)
        Case "\mid"
            LookupSymbol = ChrW(&H2223)
        Case "\prime"
            LookupSymbol = ChrW(&H2032)
        Case "\sim"
            LookupSymbol = ChrW(&H223C)
        Case "\forall"
            LookupSymbol = ChrW(&H2200)
        Case "\exists"
            LookupSymbol = ChrW(&H2203)
        Case "\nexists"
            LookupSymbol = ChrW(&H2204)
        Case "\subset"
            LookupSymbol = ChrW(&H2282)
        Case "\subseteq"
            LookupSymbol = ChrW(&H2286)
        Case "\supset"
            LookupSymbol = ChrW(&H2283)
        Case "\supseteq"
            LookupSymbol = ChrW(&H2287)
        Case "\cup"
            LookupSymbol = ChrW(&H222A)
        Case "\cap"
            LookupSymbol = ChrW(&H2229)
        Case "\langle"
            LookupSymbol = ChrW(&H27E8)
        Case "\rangle"
            LookupSymbol = ChrW(&H27E9)
        Case "\lfloor"
            LookupSymbol = ChrW(&H230A)
        Case "\rfloor"
            LookupSymbol = ChrW(&H230B)
        Case "\lceil"
            LookupSymbol = ChrW(&H2308)
        Case "\rceil"
            LookupSymbol = ChrW(&H2309)
        Case "\circ"
            LookupSymbol = ChrW(&H2218)
        Case "\oplus"
            LookupSymbol = ChrW(&H2295)
        Case "\otimes"
            LookupSymbol = ChrW(&H2297)
        Case "\perp"
            LookupSymbol = ChrW(&H22A5)
        Case "\equiv"
            LookupSymbol = ChrW(&H2261)
        Case "\propto"
            LookupSymbol = ChrW(&H221D)
        Case "\cong"
            LookupSymbol = ChrW(&H2245)
        Case "\simeq"
            LookupSymbol = ChrW(&H2243)
        Case "\triangle"
            LookupSymbol = ChrW(&H25B3)
        Case "\emptyset", "\varnothing"
            LookupSymbol = ChrW(&H2205)
        Case "\setminus"
            LookupSymbol = ChrW(&H2216)
        Case Else
            LookupSymbol = ""
    End Select
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
        Case "\tilde"
            LookupAccent = ChrW(&H303)
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
    Select Case ch
        Case "+", "-", "=", "<", ">", "*", "/", ","
            IsOperatorChar = True
        Case ChrW(&H2212) ' Minus sign
            IsOperatorChar = True
        Case ChrW(&HD7), ChrW(&HF7) ' Times, Divide
            IsOperatorChar = True
        Case ChrW(&H2264), ChrW(&H2265), ChrW(&H2260), ChrW(&H2248) ' le, ge, ne, approx
            IsOperatorChar = True
        Case ChrW(&H2208), ChrW(&H2209), ChrW(&H2192) ' in, notin, rightarrow
            IsOperatorChar = True
        Case ChrW(&H2223), ChrW(&H2225), ChrW(&H223C) ' mid, parallel, sim
            IsOperatorChar = True
        Case ChrW(&H2218), ChrW(&H2295), ChrW(&H2297), ChrW(&H22A5) ' circ, oplus, otimes, perp
            IsOperatorChar = True
        Case ChrW(&H2261), ChrW(&H221D), ChrW(&H2245), ChrW(&H2243), ChrW(&H2216) ' equiv, propto, cong, simeq, setminus
            IsOperatorChar = True
        Case ChrW(&H222A), ChrW(&H2229), ChrW(&H2282), ChrW(&H2286), ChrW(&H2283), ChrW(&H2287) ' cup, cap, sub, sube, sup, supe
            IsOperatorChar = True
        Case ChrW(&H2262), ChrW(&H2284), ChrW(&H2288), ChrW(&H2249), ChrW(&H2241) ' nequiv, nsub, nsube, napprox, nsim
            IsOperatorChar = True
        Case ChrW(&H226E), ChrW(&H226F), ChrW(&H2270), ChrW(&H2271) ' nless, ngtr, nleq, ngeq
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
    Dim res As String
    res = s
    res = Replace(res, "&", "&amp;")
    res = Replace(res, "<", "&lt;")
    res = Replace(res, ">", "&gt;")
    res = Replace(res, """", "&quot;")
    res = Replace(res, "'", "&apos;")
    EscapeXml = res
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
