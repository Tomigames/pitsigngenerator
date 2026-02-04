Option Explicit

Sub BuildPitSigns_TopToBottom()
    Dim templateIndex As Long
    Dim startRow As Long, endRow As Long
    Dim templateSlide As Slide, dupRange As SlideRange
    Dim templateID As Long
    Dim r As Long, idx As Long
    Dim insertPos As Long

    ' --- Prompts ---
    templateIndex = CLng(InputBox("Template slide number (e.g., 1):"))
    startRow = CLng(InputBox("Excel start row (e.g., 2):"))

    ' --- Connect to Excel (must already be open) ---
    Dim xlApp As Object, xlWB As Object, xlWS As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo 0

    If xlApp Is Nothing Then
        MsgBox "Open Excel first (with team list in columns A and B).", vbExclamation
        Exit Sub
    End If

    Set xlWB = xlApp.ActiveWorkbook
    Set xlWS = xlApp.ActiveSheet

    ' Auto-detect last row in column A
    endRow = xlWS.Cells(xlWS.Rows.Count, 1).End(-4162).Row  ' -4162 = xlUp

    ' --- Lock the template slide by ID (so deletions won't break indexing) ---
    Set templateSlide = ActivePresentation.Slides(templateIndex)
    templateID = templateSlide.SlideID

    ' --- Delete all slides except the template (safe) ---
    For idx = ActivePresentation.Slides.Count To 1 Step -1
        If ActivePresentation.Slides(idx).SlideID <> templateID Then
            ActivePresentation.Slides(idx).Delete
        End If
    Next idx

    ' After deletions, template is now slide 1
    Set templateSlide = ActivePresentation.Slides(1)
    insertPos = 1   ' we will insert new slides starting after slide 1

    ' --- Build slides TOP -> BOTTOM, inserting in order ---
    For r = startRow To endRow
        Dim teamNum As String, teamName As String
        teamNum = Trim(CStr(xlWS.Cells(r, 1).Text))
        teamName = Trim(CStr(xlWS.Cells(r, 2).Text))

        If Len(teamNum) = 0 And Len(teamName) = 0 Then
            ' skip blank rows
        Else
            ' Duplicate the template
            Set dupRange = templateSlide.Duplicate()

            ' Move the newly duplicated slide directly after the last inserted slide
            insertPos = insertPos + 1
            dupRange(1).MoveTo insertPos

            ' Replace placeholders on that new slide
            ReplaceTextEverywhere dupRange(1), "{{NUM}}", teamNum
            ReplaceTextEverywhere dupRange(1), "{{NAME}}", teamName
        End If
    Next r

    ' OPTIONAL: delete the template slide (leave commented unless you want it gone)
    ' ActivePresentation.Slides(1).Delete

    MsgBox "Done. Built signs from row " & startRow & " to " & endRow & " (top-to-bottom).", vbInformation
End Sub

Private Sub ReplaceTextEverywhere(ByVal sld As Slide, ByVal findTxt As String, ByVal replTxt As String)
    Dim shp As Shape
    For Each shp In sld.Shapes
        ReplaceInShape shp, findTxt, replTxt
    Next shp
End Sub

Private Sub ReplaceInShape(ByVal shp As Shape, ByVal findTxt As String, ByVal replTxt As String)
    Dim i As Long

    ' Groups: recurse
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            ReplaceInShape shp.GroupItems(i), findTxt, replTxt
        Next i
        Exit Sub
    End If

    ' Placeholders (Title/Content/etc.)
    If shp.Type = msoPlaceholder Then
        On Error Resume Next
        If shp.HasTextFrame Then
            If shp.TextFrame.HasText Then
                shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, findTxt, replTxt)
            End If
        End If
        On Error GoTo 0
        Exit Sub
    End If

    ' Normal text frames
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            shp.TextFrame.TextRange.Text = Replace(shp.TextFrame.TextRange.Text, findTxt, replTxt)
        End If
    End If

    ' Tables (if used)
    If shp.HasTable Then
        Dim rr As Long, cc As Long
        For rr = 1 To shp.Table.Rows.Count
            For cc = 1 To shp.Table.Columns.Count
                shp.Table.Cell(rr, cc).Shape.TextFrame.TextRange.Text = _
                    Replace(shp.Table.Cell(rr, cc).Shape.TextFrame.TextRange.Text, findTxt, replTxt)
            Next cc
        Next rr
    End If
End Sub
