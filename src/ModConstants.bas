Attribute VB_Name = "ModConstants"
Option Explicit
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.

' ==============================================================================
' Module: ModConstants
' Purpose: Core constants, enumerations, Unicode string decoder, and regex patterns
' Architecture: Academic Document Format Engine
' ==============================================================================

' --- System Bookmark Identifiers ---
Public Const BM_PREFIX_TOC As String = "bm_toc_"
Public Const BM_PREFIX_TBL As String = "bm_tbl_"
Public Const BM_PREFIX_FIG As String = "bm_fig_"
Public Const BM_PREFIX_REF As String = "bm_ref_"
Public Const BM_SYS_GENERATED_BLOCK As String = "bm_sys_generated_block"
Public Const BM_ZONE_TOC As String = "bm_zone_toc"
Public Const BM_ZONE_LOT As String = "bm_zone_lot"
Public Const BM_ZONE_LOF As String = "bm_zone_lof"
Public Const BM_ZONE_LOA As String = "bm_zone_loa"
Public Const BM_TOC_HEADER As String = "bm_toc_header"
Public Const BM_LOT_HEADER As String = "bm_lot_header"
Public Const BM_LOF_HEADER As String = "bm_lof_header"
Public Const BM_LOA_HEADER As String = "bm_loa_header"

' --- Engine Execution Modes ---
Public Enum EngineExecutionMode
    EngineModeFull = 1
    EngineModeListsOnly = 2
End Enum

' --- Global Silent Errors Counter ---
Public g_SilentErrors As Long

' --- Page Sync Modes ---
Public Const PageSyncRealFooter As Long = 1
Public Const PageSyncPhysicalPage As Long = 3

' --- System Markers (Case-Insensitive) ---
Public Const MARKER_CONTENT_START As String = "=contentStart()"
Public Const MARKER_CONTENT_END As String = "=contentEnd()"
Public Const MARKER_LIST_CONTENT As String = "=listContent()"
Public Const MARKER_LIST_TABLE As String = "=listTableHere()"
Public Const MARKER_LIST_FIGURE As String = "=listFigureHere()"
Public Const MARKER_LIST_ACRONYMS As String = "=listAcronyms()"

' --- Typography Defaults ---
Public Const DEFAULT_FONT_NAME As String = "Times New Roman"
Public Const DEFAULT_FONT_SIZE As Single = 13#
Public Const DUMMY_PAGE_STRING As String = "999"

' --- Target Element Classification Enumeration ---
Public Enum TargetElementType
    ElementTypeUnknown = 0
    ElementTypeLevel1 = 1
    ElementTypeLevel2 = 2
    ElementTypeLevel3 = 3
    ElementTypeLevel4 = 4
    ElementTypeTableTitle = 5
    ElementTypeFigureTitle = 6
    ElementTypeSourceNote = 7
    ElementTypeNormalParagraph = 8
    ElementTypeInlinePicture = 9
    ElementTypeLatexBlock = 10
    ElementTypeWordMath = 11
    ElementTypeList = 12
    ElementTypeTableCell = 13
    ElementTypeIgnoredManualList = 14
    ElementTypeOldToc = 15
End Enum

' --- Unicode Decoder (Prevent VBE Mojibake) ---

Private Function IsHex4(ByVal h As String) As Boolean
    Dim k As Long, code As Long
    If Len(h) <> 4 Then
        IsHex4 = False
        Exit Function
    End If
    For k = 1 To 4
        code = AscW(Mid$(h, k, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 70) Or (code >= 97 And code <= 102)) Then
            IsHex4 = False
            Exit Function
        End If
    Next k
    IsHex4 = True
End Function

Public Function Uni(ByVal s As String) As String
    ' Decodes ASCII \uXXXX sequences into UTF-16 Unicode strings
    Dim uPos As Long, startPos As Long, sLen As Long
    Dim res As String, hexCode As String
    uPos = InStr(1, s, "\u", vbBinaryCompare)
    If uPos = 0 Then
        Uni = s
        Exit Function
    End If
    sLen = Len(s)
    startPos = 1
    Do While uPos > 0
        If uPos > startPos Then
            res = res & Mid$(s, startPos, uPos - startPos)
        End If
        If uPos + 5 <= sLen Then
            hexCode = Mid$(s, uPos + 2, 4)
            If IsHex4(hexCode) Then
                res = res & ChrW(CLng("&H" & hexCode))
                startPos = uPos + 6
            Else
                res = res & "\u"
                startPos = uPos + 2
            End If
        Else
            res = res & Mid$(s, uPos)
            startPos = sLen + 1
            Exit Do
        End If
        If startPos > sLen Then
            Exit Do
        End If
        uPos = InStr(startPos, s, "\u", vbBinaryCompare)
    Loop
    If startPos <= sLen Then
        res = res & Mid$(s, startPos)
    End If
    Uni = res
End Function

' --- Regex Patterns (Structural Levels & Markers) ---

Public Function GetPatternLevel2() As String
    GetPatternLevel2 = "^(\d+\.\d+\.)\s+(.+)$"
End Function
Public Function GetPatternLevel3() As String
    GetPatternLevel3 = "^(\d+\.\d+\.\d+\.)\s+(.+)$"
End Function
Public Function GetPatternLevel4() As String
    GetPatternLevel4 = "^(\d+\.\d+\.\d+\.\d+\.)\s+(.+)$"
End Function
Public Function GetPatternMarker() As String
    GetPatternMarker = "^=(contentStart|contentEnd|listContent|listTableHere|listFigureHere|listAcronyms)\(\)\s*$"
End Function

' --- Measurement Utilities ---

Public Function CmToPoints(ByVal cmValue As Single) As Single
    CmToPoints = cmValue * 72# / 2.54#
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
