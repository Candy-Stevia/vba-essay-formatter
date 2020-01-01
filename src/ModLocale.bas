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
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = False
    regEx.Global = False
    regEx.Pattern = "^(X{1,3}(?:IX|IV|V?I{0,3})|IX|IV|V?I{1,3}|V)[\.:]\s+\S.*$"
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
    GetPatternSourceNote = "^(Source|" & GetViNguon() & "|Note|" & GetViGhiChu() & "):\s*(.*)$"
End Function

Public Function GetPrefixCapPattern() As String
    Dim tblFig As String, srcNote As String
    tblFig = "(?:Table|" & ModConstants.Uni("B(?:\u1ea3|a)ng") & "|" & GetFigureKeywords() & ")\s+" & NumToken() & CaptionSep()
    srcNote = "(?:Source|" & GetViNguon() & "|Note|" & GetViGhiChu() & "):\s*"
    GetPrefixCapPattern = "^(?:" & tblFig & "|" & srcNote & "|(?:\d+\.)+\s*)"
End Function

Public Function IsPrefixOnlyHeading(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    
    Dim enPat As String, viPat As String
    enPat = "(Chapter\s+[A-Za-z0-9]+[:\.]|Appendix(\s+[A-Za-z0-9]+)?[:\.]|" & _
            "\d+(\.\d+)*\.|" & _
            "Table\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*\.|" & _
            "(Figure|Fig\.?|Chart|Diagram)\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*[:\.]|" & _
            "Source:|Note:)"
            
    viPat = ModConstants.Uni("(Ch\u01b0\u01a1ng\s+[A-Za-z0-9]+[:\.]|Ph\u1ee5 l\u1ee5c(\s+[A-Za-z0-9]+)?[:\.]|" & _
            "B(?:\u1ea3|a)ng\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*\.|" & _
            "(H(?:\u00ec|i)nh(?:\s+(?:\u1ea3nh|v\u1ebd))?|S\u01a1 \u0111\u1ed3|Bi\u1ec3u \u0111\u1ed3|\u0110\u1ed3 th\u1ecb)\s+[A-Za-z0-9]+(?:[\.\-_][A-Za-z0-9]+)*[:\.]|" & _
            "Ngu\u1ed1n:|Ghi ch\u00fa:)")
            
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Global = False
    regEx.Pattern = "^(" & enPat & "|" & viPat & ")$"
    IsPrefixOnlyHeading = regEx.Test(t)
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
