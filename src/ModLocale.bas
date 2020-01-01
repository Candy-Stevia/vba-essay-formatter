Attribute VB_Name = "ModLocale"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModLocale
' Purpose: Bilingual dictionary, Unicode keywords, and regex pattern providers
' Architecture: Academic Document Format Engine (Full English & Vietnamese Support)
' ==============================================================================

Public Const LANG_VI As String = "lang_vi"
Public Const LANG_EN As String = "lang_en"

Public Function GetViBang() As String
    GetViBang = "B" & ChrW(7843) & "ng"
End Function

Public Function GetViHinh() As String
    GetViHinh = "H" & ChrW(236) & "nh"
End Function

Public Function GetViNguon() As String
    GetViNguon = "Ngu" & ChrW(7891) & "n"
End Function

Public Function GetViGhiChu() As String
    GetViGhiChu = "Ghi ch" & ChrW(250)
End Function

Public Function GetTocHeaderTitle(ByVal lang As String) As String
    If lang = LANG_VI Then GetTocHeaderTitle = ModConstants.Uni("M\u1ee4C L\u1ee4C") Else GetTocHeaderTitle = "TABLE OF CONTENTS"
End Function

Public Function GetLotHeaderTitle(ByVal lang As String) As String
    If lang = LANG_VI Then GetLotHeaderTitle = ModConstants.Uni("DANH M\u1ee4C B\u1ea2NG BI\u1ec2U") Else GetLotHeaderTitle = "LIST OF TABLES"
End Function

Public Function GetLofHeaderTitle(ByVal lang As String) As String
    If lang = LANG_VI Then GetLofHeaderTitle = ModConstants.Uni("DANH M\u1ee4C H\u00ccNH \u1ea2NH") Else GetLofHeaderTitle = "LIST OF FIGURES"
End Function

Public Function GetLoaHeaderTitle(ByVal lang As String) As String
    If lang = LANG_VI Then GetLoaHeaderTitle = ModConstants.Uni("DANH M\u1ee4C T\u1eea VI\u1ebeT T\u1eaeT") Else GetLoaHeaderTitle = "LIST OF ACRONYMS"
End Function

Public Function GetFigureKeywords() As String
    Dim viFigPrefix As String
    viFigPrefix = "H(?:\u00ec|i)nh(?:\s+(?:\u1ea3nh|v\u1ebd))?|S\u01a1 \u0111\u1ed3|Bi\u1ec3u \u0111\u1ed3|\u0110\u1ed3 th\u1ecb"
    GetFigureKeywords = "Figure|Fig\.?|Chart|Diagram|" & ModConstants.Uni(viFigPrefix)
End Function

Public Function NumToken() As String
    NumToken = "[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*"
End Function

Public Function CaptionSep() As String
    ' Dau ':' (khoang trang tuy y) | Dau '.' bat buoc >=1 khoang trang | Dau gach (-, en/em dash) bat buoc khoang trang 2 ben
    CaptionSep = "(?:\s*:\s*|\.\s+|\s+[\-" & ModConstants.Uni("\u2013\u2014") & "]\s+)"
End Function

Public Function GetPatternLevel1() As String
    Dim enParts As String, viParts As String
    enParts = "acknowledg(e)?ments?|(list of )?acronyms?|(list of )?abbreviations?|" & _
              "declaration( of originality)?|abstract|introduction|conclusions?|" & _
              "references?|bibliography|appendices|appendixes|" & _
              "chapter\s+[A-Za-z0-9]+[:\.]\s*.+|appendix(\s+[A-Za-z0-9]+)?([:\.]\s*.+)?"
    
    viParts = ModConstants.Uni("l\u1eddi c[\u1ea3\u00e1]m \u01a1n|l\u1eddi cam \u0111oan|t\u00f3m t\u1eaft|" & _
                               "(l\u1eddi )?m\u1edf \u0111\u1ea7u|gi\u1edbi thi\u1ec7u|" & _
                               "k\u1ebft lu\u1eadn( v\u00e0 ki\u1ebfn ngh\u1ecb)?|" & _
                               "(danh m\u1ee5c )?t\u00e0i li\u1ec7u tham kh\u1ea3o|" & _
                               "(danh m\u1ee5c )?(c\u00e1c )?(t\u1eeb|ch\u1eef) vi\u1ebft t\u1eaft|" & _
                               "danh m\u1ee5c k\u00fd hi\u1ec7u v\u00e0 ch\u1eef vi\u1ebft t\u1eaft|" & _
                               "ch\u01b0\u01a1ng\s+[A-Za-z0-9]+[:\.]\s*.+|ph\u1ea7n\s+[A-Za-z0-9]+[:\.]\s*.+|ph\u1ee5 l\u1ee5c(\s+[A-Za-z0-9]+)?([:\.]\s*.+)?")
    
    GetPatternLevel1 = "^(" & enParts & "|" & viParts & ")$"
End Function

Public Function IsRomanLevel1(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) < 4 Then Exit Function
    Dim ch1 As String
    ch1 = Left$(t, 1)
    If ch1 <> "I" And ch1 <> "V" And ch1 <> "X" Then Exit Function
    Static regEx As Object
    If regEx Is Nothing Then
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.IgnoreCase = False
        regEx.Global = False
        regEx.Pattern = "^(X{1,3}(?:IX|IV|V?I{0,3})|IX|IV|V?I{1,3}|V)[\.:]\s+\S.*$"
    End If
    IsRomanLevel1 = regEx.Test(t)
End Function

Public Function GetPatternManualList() As String
    Dim enParts As String, viParts As String
    enParts = "tables? of contents?|list of tables?|list of figures?"
    viParts = ModConstants.Uni("m\u1ee5c l\u1ee5c|(danh m\u1ee5c|danh s\u00e1ch) b\u1ea3ng( bi\u1ec3u)?|" & _
                               "(danh m\u1ee5c|danh s\u00e1ch) h\u00ecnh( \u1ea3nh| v\u1ebd)?|danh m\u1ee5c s\u01a1 \u0111\u1ed3")
    GetPatternManualList = "^(" & enParts & "|" & viParts & ")$"
End Function

Public Function GetPatternManualToc() As String
    GetPatternManualToc = "^(tables? of contents?|" & ModConstants.Uni("m\u1ee5c l\u1ee5c") & ")$"
End Function

Public Function GetPatternManualLot() As String
    GetPatternManualLot = "^(list of tables?|" & ModConstants.Uni("(danh m\u1ee5c|danh s\u00e1ch) b\u1ea3ng( bi\u1ec3u)?") & ")$"
End Function

Public Function GetPatternManualLof() As String
    GetPatternManualLof = "^(list of figures?|" & ModConstants.Uni("(danh m\u1ee5c|danh s\u00e1ch) h\u00ecnh( \u1ea3nh| v\u1ebd)?|danh m\u1ee5c s\u01a1 \u0111\u1ed3") & ")$"
End Function

Public Function GetPatternTable() As String
    GetPatternTable = "^(Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & ")\s+(" & NumToken() & ")" & CaptionSep() & "(\S.*)$"
End Function

Public Function GetPatternFigure() As String
    GetPatternFigure = "^(" & GetFigureKeywords() & ")\s+(" & NumToken() & ")" & CaptionSep() & "(\S.*)$"
End Function

Public Function GetPatternSourceNote() As String
    GetPatternSourceNote = "^(Source|" & GetViNguon() & "|Note|" & GetViGhiChu() & ")\s*:\s*(.*)$"
End Function

Public Function GetPrefixCapPattern() As String
    Dim tblFig As String, srcNote As String
    tblFig = "(?:Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & "|" & GetFigureKeywords() & ")\s+" & NumToken() & CaptionSep()
    srcNote = "(?:Source|" & GetViNguon() & "|Note|" & GetViGhiChu() & ")\s*:\s*"
    GetPrefixCapPattern = "^\s*(?:" & tblFig & "|" & srcNote & "|(?:\d+\.)+\s*)"
End Function

Public Function IsPrefixOnlyHeading(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    
    Dim sep As String
    sep = "(?:\s*:|\.|\s+[\-" & ModConstants.Uni("\u2013\u2014") & "])"
    
    Dim ch1 As String
    ch1 = Left$(t, 1)
    If ch1 = "I" Or ch1 = "V" Or ch1 = "X" Then
        Static regRoman As Object
        If regRoman Is Nothing Then
            Set regRoman = CreateObject("VBScript.RegExp")
            regRoman.IgnoreCase = False
            regRoman.Global = False
            regRoman.Pattern = "^(X{1,3}(?:IX|IV|V?I{0,3})|IX|IV|V?I{1,3}|V)" & sep & "$"
        End If
        If regRoman.Test(t) Then
            IsPrefixOnlyHeading = True
            Exit Function
        End If
    End If
    
    Static regEx As Object
    If regEx Is Nothing Then
        Dim enPat As String, viPat As String
        enPat = "(Chapter\s+[A-Za-z0-9]+" & sep & "|Appendix(\s+[A-Za-z0-9]+)?" & sep & "|" & _
                "\d+(\.\d+)*\.|" & _
                "Table\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*" & sep & "|" & _
                "(Figure|Fig\.?|Chart|Diagram)\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*" & sep & "|" & _
                "Source\s*:|Note\s*:)"
                
        viPat = "(" & ModConstants.Uni("Ch\u01b0\u01a1ng") & "\s+[A-Za-z0-9]+" & sep & "|" & _
                ModConstants.Uni("Ph\u1ea7n") & "\s+[A-Za-z0-9]+" & sep & "|" & _
                ModConstants.Uni("Ph\u1ee5 l\u1ee5c") & "(\s+[A-Za-z0-9]+)?" & sep & "|" & _
                ModConstants.Uni("B(?:\u1ea3|a)ng") & "\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*" & sep & "|" & _
                ModConstants.Uni("(H(?:\u00ec|i)nh(?:\s+(?:\u1ea3nh|v\u1ebd))?|S\u01a1 \u0111\u1ed3|Bi\u1ec3u \u0111\u1ed3|\u0110\u1ed3 th\u1ecb)") & _
                "\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*" & sep & "|" & _
                ModConstants.Uni("Ngu\u1ed3n\s*:|Ghi ch\u00fa\s*:") & ")"
                
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.IgnoreCase = True
        regEx.Global = False
        regEx.Pattern = "^(" & enPat & "|" & viPat & ")$"
    End If
    IsPrefixOnlyHeading = regEx.Test(t)
End Function

Public Function ClassifyRefZoneHeading(ByVal cleanTxt As String, Optional ByVal exactEnter As Boolean = False) As Long
    Dim t As String
    t = Trim$(Replace(Replace(Replace(cleanTxt, vbCr, ""), Chr(11), " "), Chr(7), ""))
    If Len(t) = 0 Then Exit Function
    Static reStrip As Object, reStripRoman As Object, reEnter As Object, reEnterExact As Object, reExit As Object
    If reStrip Is Nothing Then
        Set reStrip = CreateObject("VBScript.RegExp")
        reStrip.IgnoreCase = True
        reStrip.Global = False
        reStrip.Pattern = "^(?:(?:chapter|" & ModConstants.Uni("ch\u01b0\u01a1ng") & "|part|" & _
                          ModConstants.Uni("ph\u1ea7n") & ")\s+[a-z0-9]+\s*[:\.\-]?\s*|\d+(?:\.\d+)*\s*[:\.\-]?\s+)"
        Set reStripRoman = CreateObject("VBScript.RegExp")
        reStripRoman.IgnoreCase = False
        reStripRoman.Global = False
        reStripRoman.Pattern = "^[IVXLCDM]+\s*[:\.\-]\s+"
        Dim enterBase As String
        enterBase = "(?:references?|bibliography|tai lieu tham khao|danh muc tai lieu tham khao|" & _
                    ModConstants.Uni("t\u00e0i li\u1ec7u tham kh\u1ea3o|danh m\u1ee5c t\u00e0i li\u1ec7u tham kh\u1ea3o") & ")"
        Set reEnter = CreateObject("VBScript.RegExp")
        reEnter.IgnoreCase = True
        reEnter.Pattern = "^" & enterBase & "(?:$|\s|[:\.\-])"
        Set reEnterExact = CreateObject("VBScript.RegExp")
        reEnterExact.IgnoreCase = True
        reEnterExact.Pattern = "^" & enterBase & "$"
        Set reExit = CreateObject("VBScript.RegExp")
        reExit.IgnoreCase = True
        reExit.Pattern = "^(?:appendices|appendix|glossary|phu luc|" & ModConstants.Uni("ph\u1ee5 l\u1ee5c") & ")(?:$|\s|[:\.\-])"
    End If
    t = Trim$(reStrip.Replace(t, ""))
    t = Trim$(reStripRoman.Replace(t, ""))
    If Len(t) = 0 Then Exit Function
    If exactEnter Then
        If reEnterExact.Test(t) Then
            ClassifyRefZoneHeading = 1
        ElseIf reExit.Test(t) Then
            ClassifyRefZoneHeading = -1
        End If
    Else
        If reEnter.Test(t) Then
            ClassifyRefZoneHeading = 1
        ElseIf reExit.Test(t) Then
            ClassifyRefZoneHeading = -1
        End If
    End If
End Function

Private Function TestSingleCurrencyUnit(ByVal rawUnit As String, ByVal numPart As String, ByVal reCurr As Object) As Boolean
    If rawUnit = "M" Or rawUnit = "B" Or rawUnit = "K" Then
        TestSingleCurrencyUnit = True
    ElseIf rawUnit = "k" Then
        If Val(Replace(Replace(numPart, ",", ""), " ", "")) >= 10# Then
            TestSingleCurrencyUnit = True
        End If
    ElseIf reCurr.Test(rawUnit) Then
        TestSingleCurrencyUnit = True
    End If
End Function

Public Function IsCurrencyUnitMatch(ByVal rawUnitPart As String, ByVal numPart As String) As Boolean
    Static reCurr As Object, rePeriod As Object
    If reCurr Is Nothing Then
        Set reCurr = CreateObject("VBScript.RegExp")
        reCurr.IgnoreCase = True
        reCurr.Global = False
        reCurr.Pattern = "^(?:usd|vnd|eur|gbp|jpy|cny|chf|tr|bn|dong|nghin|ngan|trieu|ty|dollar|dollars|cent|cents|thousand|million|billion|" & _
                         ModConstants.Uni("tri\u1ec7u|t\u1ef7|ngh\u00ecn|ng\u00e0n|\u0111\u1ed3ng|\u0111") & ")$"
        Set rePeriod = CreateObject("VBScript.RegExp")
        rePeriod.IgnoreCase = True
        rePeriod.Global = False
        rePeriod.Pattern = "^(?:month|mo|year|yr|day|ngay|unit|thang|nam|nguoi|don vi|" & _
                           ModConstants.Uni("th\u00e1ng|n\u0103m|ng\u00e0y|ng\u01b0\u1eddi|\u0111\u01a1n v\u1ecb") & ")$"
    End If
    Dim slashPos As Long, leftTok As String, rightTok As String
    slashPos = InStr(1, rawUnitPart, "/")
    If slashPos > 0 Then
        leftTok = Trim$(Left$(rawUnitPart, slashPos - 1))
        rightTok = Trim$(Mid$(rawUnitPart, slashPos + 1))
        If Len(rightTok) > 0 And rePeriod.Test(rightTok) Then
            If Len(leftTok) = 0 Then
                IsCurrencyUnitMatch = True
            Else
                IsCurrencyUnitMatch = TestSingleCurrencyUnit(leftTok, numPart, reCurr)
            End If
        End If
    Else
        IsCurrencyUnitMatch = TestSingleCurrencyUnit(rawUnitPart, numPart, reCurr)
    End If
End Function
Public Function IsRefSubheading(ByVal cleanTxt As String) As Boolean
    Dim t As String
    t = Trim$(cleanTxt)
    If Len(t) = 0 Or Len(t) > 100 Then Exit Function
    Static reStrip As Object, reSub As Object
    If reSub Is Nothing Then
        Set reStrip = CreateObject("VBScript.RegExp")
        reStrip.IgnoreCase = True
        reStrip.Global = False
        reStrip.Pattern = "^(?:\d+(?:\.\d+)*\.?|[a-zA-Z]\)|[IVXLCDM]+[\.:])\s+"
        Set reSub = CreateObject("VBScript.RegExp")
        reSub.IgnoreCase = True
        reSub.Global = False
        Dim viSub As String, enSub As String
        viSub = ModConstants.Uni("t\u00e0i li\u1ec7u (?:tham kh\u1ea3o )?(?:ti\u1ebfng vi\u1ec7t|ti\u1ebfng anh|n\u01b0\u1edbc ngo\u00e0i)|" & _
                                 "v\u0103n b\u1ea3n ph\u00e1p lu\u1eadt|ngu\u1ed3n (?:internet|tr\u1ef1c tuy\u1ebfn)")
        enSub = "vietnamese references|english references|primary sources|secondary sources|books|journal articles|online sources"
        reSub.Pattern = "^(?:" & viSub & "|" & enSub & ")[\s:\.\-\u2013\u2014]*$"
    End If
    t = Trim$(reStrip.Replace(t, ""))
    IsRefSubheading = reSub.Test(t)
End Function

' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
