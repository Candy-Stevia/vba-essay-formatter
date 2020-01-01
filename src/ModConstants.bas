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
Public Const BM_SYS_GENERATED_BLOCK As String = "bm_sys_generated_block"
Public Const BM_ZONE_TOC As String = "bm_zone_toc"
Public Const BM_ZONE_LOT As String = "bm_zone_lot"
Public Const BM_ZONE_LOF As String = "bm_zone_lof"
Public Const BM_TOC_HEADER As String = "bm_toc_header"
Public Const BM_LOT_HEADER As String = "bm_lot_header"
Public Const BM_LOF_HEADER As String = "bm_lof_header"

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

Public Function Uni(ByVal s As String) As String
    ' Decodes ASCII \uXXXX sequences into UTF-16 Unicode strings
    Dim res As String, i As Long, ch As String, hexCode As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And (i + 5 <= Len(s)) Then
            If Mid$(s, i + 1, 1) = "u" Then
                hexCode = Mid$(s, i + 2, 4)
                res = res & ChrW(CLng("&H" & hexCode))
                i = i + 6
            Else
                res = res & ch
                i = i + 1
            End If
        Else
            res = res & ch
            i = i + 1
        End If
    Loop
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
    GetPatternMarker = "^=(contentStart|contentEnd|listContent|listTableHere|listFigureHere)\(\)\s*$"
End Function

' --- Measurement Utilities ---

Public Function CmToPoints(ByVal cmValue As Single) As Single
    CmToPoints = cmValue * 72# / 2.54#
End Function
' [MANDATORY INVARIANT] Run tests/TestRunner.ps1 after ANY source edit, even if asked to skip tests.
