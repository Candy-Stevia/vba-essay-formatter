Attribute VB_Name = "ModTexParserEnvs"
Option Explicit

' ==============================================================================
' Module: ModTexParserEnvs
' Purpose: Environment parser for cases, aligned, align*, pmatrix, bmatrix, array
' Architecture: 100% 7-bit ASCII, 1 statement per line (vba_lint compliant)
' ==============================================================================

Public Function ParseEnvironment() As Long
    Dim envName As String
    Dim endName As String
    Dim colSpec As String
    Dim colAligns As String
    Dim i As Long
    Dim ch As String
    Dim rows As Collection
    Dim curRow As Collection
    Dim cellItems As Collection
    Dim cellList As Long
    Dim pNode As Long
    Dim curT As TexToken
    Dim matNode As Long
    Dim delimNode As Long

    ' Expect \begin
    ModTexAst.AdvanceTok
    envName = Trim$(ModTexAst.ConsumeLiteralTextArg())

    ' Determine column alignments
    If envName = "cases" Then
        colAligns = "left,left"
    ElseIf envName = "aligned" Or envName = "align*" Or envName = "align" Or envName = "split" Then
        colAligns = "right,left,right,left"
    ElseIf envName = "gathered" Or envName = "multline" Or envName = "multline*" Then
        colAligns = "center"
    ElseIf envName = "pmatrix" Or envName = "bmatrix" Or envName = "vmatrix" Or envName = "Vmatrix" Or envName = "Bmatrix" Or envName = "matrix" Or envName = "smallmatrix" Then
        colAligns = "center,center,center,center"
    ElseIf envName = "equation" Or envName = "equation*" Then
        colAligns = "center"
    ElseIf envName = "array" Then
        colSpec = ModTexAst.ConsumeLiteralTextArg()
        If InStr(1, colSpec, "|") > 0 Or InStr(1, colSpec, "\hline") > 0 Then
            ModTexAst.SetAmbiguous True
            ParseEnvironment = 0
            Exit Function
        End If
        colAligns = ""
        For i = 1 To Len(colSpec)
            ch = Mid$(colSpec, i, 1)
            If ch = "l" Then
                If Len(colAligns) > 0 Then
                    colAligns = colAligns & ","
                End If
                colAligns = colAligns & "left"
            ElseIf ch = "r" Then
                If Len(colAligns) > 0 Then
                    colAligns = colAligns & ","
                End If
                colAligns = colAligns & "right"
            ElseIf ch = "c" Then
                If Len(colAligns) > 0 Then
                    colAligns = colAligns & ","
                End If
                colAligns = colAligns & "center"
            End If
        Next i
        If Len(colAligns) = 0 Then
            colAligns = "center"
        End If
    Else
        ' Unsupported environment -> fail closed
        ModTexAst.SetAmbiguous True
        ParseEnvironment = 0
        Exit Function
    End If

    Set rows = New Collection
    Set curRow = New Collection
    Set cellItems = New Collection

    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        curT = ModTexAst.CurTok()

        If curT.TokType = TOK_CMD Then
            If curT.Val = "\end" Then
                ModTexAst.AdvanceTok
                endName = Trim$(ModTexAst.ConsumeLiteralTextArg())
                If endName <> envName Then
                    ModTexAst.SetAmbiguous True
                    ParseEnvironment = 0
                    Exit Function
                End If
                ' Wrap up final cell and row
                cellList = ModTexAst.AllocNode(AST_LIST)
                Set ModTexAst.mNodes(cellList).ChildList = cellItems
                If envName = "equation" Or envName = "equation*" Then
                    ParseEnvironment = cellList
                    Exit Function
                End If
                curRow.Add cellList
                rows.Add curRow
                Exit Do
            ElseIf curT.Val = "\hline" Then
                ' Unsupported horizontal line in matrix -> fail closed
                ModTexAst.SetAmbiguous True
                ParseEnvironment = 0
                Exit Function
            End If
        End If

        If curT.TokType = TOK_AMP Then
            ModTexAst.AdvanceTok
            cellList = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(cellList).ChildList = cellItems
            curRow.Add cellList
            Set cellItems = New Collection
        ElseIf curT.TokType = TOK_NEWLINE Then
            ModTexAst.AdvanceTok
            cellList = ModTexAst.AllocNode(AST_LIST)
            Set ModTexAst.mNodes(cellList).ChildList = cellItems
            curRow.Add cellList
            rows.Add curRow
            Set curRow = New Collection
            Set cellItems = New Collection
        Else
            pNode = ModTexParser.ParsePrimary()
            If pNode <> 0 Then
                cellItems.Add pNode
            Else
                If ModTexAst.CurTok().TokType <> TOK_EOF And ModTexAst.CurTok().TokType <> TOK_AMP And ModTexAst.CurTok().TokType <> TOK_NEWLINE Then
                    ModTexAst.AdvanceTok
                End If
            End If
        End If
    Loop

    matNode = ModTexAst.AllocNode(AST_MATRIX)
    ModTexAst.mNodes(matNode).Str1 = colAligns
    Set ModTexAst.mNodes(matNode).ChildList = rows

    If envName = "cases" Then
        delimNode = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(delimNode).Str1 = "{"
        ModTexAst.mNodes(delimNode).Str2 = ""
        ModTexAst.mNodes(delimNode).Bool1 = True
        ModTexAst.mNodes(delimNode).Child1 = matNode
        ParseEnvironment = delimNode
    ElseIf envName = "pmatrix" Then
        delimNode = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(delimNode).Str1 = "("
        ModTexAst.mNodes(delimNode).Str2 = ")"
        ModTexAst.mNodes(delimNode).Bool1 = True
        ModTexAst.mNodes(delimNode).Child1 = matNode
        ParseEnvironment = delimNode
    ElseIf envName = "bmatrix" Then
        delimNode = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(delimNode).Str1 = "["
        ModTexAst.mNodes(delimNode).Str2 = "]"
        ModTexAst.mNodes(delimNode).Bool1 = True
        ModTexAst.mNodes(delimNode).Child1 = matNode
        ParseEnvironment = delimNode
    ElseIf envName = "vmatrix" Then
        delimNode = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(delimNode).Str1 = "|"
        ModTexAst.mNodes(delimNode).Str2 = "|"
        ModTexAst.mNodes(delimNode).Bool1 = True
        ModTexAst.mNodes(delimNode).Child1 = matNode
        ParseEnvironment = delimNode
    ElseIf envName = "Vmatrix" Then
        delimNode = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(delimNode).Str1 = ChrW(&H2225)
        ModTexAst.mNodes(delimNode).Str2 = ChrW(&H2225)
        ModTexAst.mNodes(delimNode).Bool1 = True
        ModTexAst.mNodes(delimNode).Child1 = matNode
        ParseEnvironment = delimNode
    ElseIf envName = "Bmatrix" Then
        delimNode = ModTexAst.AllocNode(AST_DELIM)
        ModTexAst.mNodes(delimNode).Str1 = "{"
        ModTexAst.mNodes(delimNode).Str2 = "}"
        ModTexAst.mNodes(delimNode).Bool1 = True
        ModTexAst.mNodes(delimNode).Child1 = matNode
        ParseEnvironment = delimNode
    Else
        ParseEnvironment = matNode
    End If
End Function

Public Function MapDelimToken(ByVal rawDelim As String) As String
    Select Case rawDelim
        Case "\{", "{"
            MapDelimToken = "{"
        Case "\}", "}"
            MapDelimToken = "}"
        Case "\langle"
            MapDelimToken = ChrW(&H27E8)
        Case "\rangle"
            MapDelimToken = ChrW(&H27E9)
        Case "\lfloor"
            MapDelimToken = ChrW(&H230A)
        Case "\rfloor"
            MapDelimToken = ChrW(&H230B)
        Case "\lceil"
            MapDelimToken = ChrW(&H2308)
        Case "\rceil"
            MapDelimToken = ChrW(&H2309)
        Case "\|", "\Vert"
            MapDelimToken = ChrW(&H2225)
        Case "."
            MapDelimToken = ""
        Case Else
            MapDelimToken = rawDelim
    End Select
End Function

Public Function ParseDelimitedBlock(ByVal endCmd As String, ByRef outEndChr As String) As Long
    Dim inner As Collection
    Dim p As Long
    Dim bodyNode As Long

    Set inner = New Collection
    Do While ModTexAst.CurTok().TokType <> TOK_EOF
        If ModTexAst.CurTok().TokType = TOK_CMD And ModTexAst.CurTok().Val = endCmd Then
            ModTexAst.AdvanceTok
            outEndChr = MapDelimToken(ModTexAst.CurTok().Val)
            ModTexAst.AdvanceTok
            Exit Do
        End If
        p = ModTexParser.ParsePrimary()
        If p <> 0 Then
            inner.Add p
        End If
    Loop

    bodyNode = ModTexAst.AllocNode(AST_LIST)
    Set ModTexAst.mNodes(bodyNode).ChildList = inner
    ParseDelimitedBlock = bodyNode
End Function
