Attribute VB_Name = "main"
Option Explicit

Public Sub MakeQuiz()
    ' main
    
    ' begin
    
    On Error GoTo Finally
    
    Dim ST As Double: ST = Timer
    
    ' Specifies the template PPT file path by cell.
    Dim templatePPTFullName As String
    templatePPTFullName = Replace(ActiveSheet.Range("A1").Value, """", "")
    
    ' Gets or opens the template PPT file.
    Dim targetPresentation As Object
    Set targetPresentation = GetOrOpenPresentation(templatePPTFullName)
    If targetPresentation Is Nothing Then
        Debug.Print "Template file `" & templatePPTFullName & "` was not found."
        MsgBox "Template file `" & templatePPTFullName & "` was not found."
        GoTo Finally
    End If
    
    Dim saveFullName As String
    saveFullName = NewSaveFullNameX2P(ThisWorkbook.FullName, "_yyyy-MM-dd_HH-mm-ss", templatePPTFullName)
    
    ' Saves the template PPT file as a new file.
    ' The new file will be input into `targetPresentation` variable.
    targetPresentation.SaveAs saveFullName
    
    ' Gets the number of slides in the template PPT file.
    '`slidesCount` variable has the number of slides at this moment.
    Dim slidesCount As Long
    slidesCount = targetPresentation.Slides.Count
    
    ' Specifies the cell.
    ' Modify the range address if needed.
    Dim quizListTopLeftCell As Range
    Set quizListTopLeftCell = ThisWorkbook.ActiveSheet.Range("A3")
    
    ' Gets the range of quiz list.
    Dim quizListRange As Range
    Set quizListRange = quizListTopLeftCell.CurrentRegion
    
    '
    Dim quizList As Variant
    quizList = quizListRange.Resize(quizListRange.Rows.Count - 1).Offset(1, 0).Value
    
    '
    Dim labelRange As Range
    Set labelRange = quizListRange.Resize(1)
    
    '
    Dim labels As Variant
    labels = labelRange.Value
    
    Dim templateColumnIndex As Long
    templateColumnIndex = 1
    
    Application.StatusBar = "starting..."
    
    Debug.Print Timer - ST, "end of begin"
    
    
    ''''''''''''''''''''''''''''''
    'process
    
    Application.StatusBar = "0 / " & UBound(quizList, 1) - 1
    
    ' For all quizzes,
    '     copies template slide
    '     and pastes texts to shapes in the slides.
    Dim i As Long
    For i = 1 To UBound(quizList, 1)
        
        Dim ST2 As Double: ST2 = Timer
        
        ' Gets the template slide index from `quizList` 2D array.
        Dim templateSlideIndex As Long '
        templateSlideIndex = CLng(quizList(i, templateColumnIndex))
        
        ' Skips if `templateSlideIndex` = 0.
        If templateSlideIndex = 0 Then '
            GoTo Continue
        End If
        
        ' Skips if `templateSlideIndex` exceeds `slidesCount`.
        If templateSlideIndex > slidesCount Then
            GoTo Continue
        End If
        
        ' Gets template slide.
        ' No change will be made to `templateSlide`.
        Dim templateSlide As Object
        Set templateSlide = targetPresentation.Slides(templateSlideIndex)
        
        ' Duplicates `templateSlide` and put it in `copiesSlide`.
        ' All changes will be made to `copiedSlide` instead of `templateSlide`.
        Dim copiedSlide As Object
        Set copiedSlide = templateSlide.Duplicate
        copiedSlide.MoveTo targetPresentation.Slides.Count
        copiedSlide.SlideShowTransition.Hidden = msoFalse
        
        ' For all columns in each quiz,
        '     gets the label and
        '     pastes text to shape.
        Dim j As Long
        For j = 1 To UBound(labels, 2)
            
            If j = templateColumnIndex Then GoTo ContinueJ
            
            ' Gets the label for the column.
            Dim targetLabel As String
            targetLabel = CStr(labels(1, j))
            If targetLabel = "" Then
                GoTo ContinueJ
            End If
            
            ' Gets the text to paste to the shape.
            Dim targetText As String
            targetText = CStr(quizList(i, j))
            If targetText = "" Then
                GoTo ContinueJ
            End If
            
            ' Invokes `PasteTextToShape` procedure to paste the text to the shape.
            Call PasteTextToShape(copiedSlide, targetLabel, targetText)
            
ContinueJ:
        Next j
        
Continue:
        Debug.Print Timer - ST2, "loop " & i
        Application.StatusBar = i & " of " & UBound(quizList, 1)
        
    Next i
    
    Debug.Print Timer - ST, "end of process"
    
    
    '''''''''''''''''''''''''''
    'end
    
    targetPresentation.Save
    
    Debug.Print Timer - ST, "end"
    Debug.Print
    
    ThisWorkbook.Activate
    MsgBox "終了しました。"
    
Finally:
    If Err.Number > 0 Then MsgBox Err.Number & vbCrLf & Err.Description
    
    Application.StatusBar = False
    
End Sub


Private Function GetOrOpenPresentation(file_fullname As String) As Object
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(file_fullname) = False Then
        Debug.Print "Designated file was not found."
        Set FSO = Nothing
        Exit Function
    End If
    
    Set FSO = Nothing
    
    Dim PPT As Object
    Set PPT = CreateObject("PowerPoint.Application")
    
    '
    Dim PR As Object
    For Each PR In PPT.Presentations
        If PR.FullName = file_fullname Then
            Set GetOrOpenPresentation = PR
            Set PPT = Nothing
            Exit Function
        End If
    Next PR
    
    '
    Set GetOrOpenPresentation = PPT.Presentations.Open(file_fullname, ReadOnly:=True)
    Set PPT = Nothing
    
End Function


Private Function NewSaveFullNameX2P(file_fullname As String, datetime_format As String, template_ppt_fullname As String) As String
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    '
    Do
        Dim fileParentPath As String
        fileParentPath = FSO.GetParentFolderName(file_fullname)
        Dim fileBaseName As String
        fileBaseName = FSO.GetBaseName(file_fullname)
        Dim fileExtension As String
        fileExtension = FSO.GetExtensionName(template_ppt_fullname)
        Dim formattedDateTime As String
        formattedDateTime = Format(Now, datetime_format)
        
        Dim result As String
        result = fileParentPath & "\" & fileBaseName & formattedDateTime & "." & fileExtension
    Loop While FSO.FileExists(result)
    
    Set FSO = Nothing
    
    NewSaveFullNameX2P = result
    
End Function


Private Sub PasteTextToShape(slide_object As Object, lookup_label As String, quiz_text As String)
    Dim targetShape As Object
    Set targetShape = GetShapeByText(slide_object, lookup_label)
    
    If targetShape Is Nothing Then
        Debug.Print "Shape with Text """ & lookup_label & """ was not found."
        Exit Sub
    End If
    
    targetShape.TextFrame.TextRange.Text = quiz_text
End Sub


Private Function GetShapeByText(target_slide As Object, target_text As String) As Object
    
    Dim SHP As Object
    For Each SHP In target_slide.Shapes
        
        If LCase(SHP.TextFrame.TextRange.Text) = LCase(target_text) Then
            Set GetShapeByText = SHP
            Exit Function
        End If
        
    Next SHP
    
End Function


Public Sub AddToContextMenu()
    
    With Application.CommandBars
        
        Dim i As Long
        For i = 1 To .Count
            
            If .Item(i).Name = "Cell" Then
                
                With .Item(i).Controls.Add(Type:=msoControlPopup, Temporary:=True)
                    .BeginGroup = True
                    .Caption = "&" & ThisWorkbook.Name

                    With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                        '.Caption = "Make &Quiz"
                        .Caption = "クイズをスライドに流し込む(&Q)"
                        .OnAction = ThisWorkbook.Name & "!MakeQuiz"
                    End With
                    
                End With
                
            End If
            
        Next i
        
    End With
    
End Sub
