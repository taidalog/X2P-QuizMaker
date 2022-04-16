Attribute VB_Name = "main"
Option Explicit

Public Sub MakeQuiz()
    ' main
    
    ' begin
    
    On Error GoTo Finally
    
    ' Specifies the template PPT file path by cell.
    Dim templatePPTFullName As String
    templatePPTFullName = Replace(ActiveSheet.Range("A1").Value, """", "")
    
    Dim PPT As Object
    Set PPT = CreateObject("PowerPoint.Application")
    
    ' Gets or opens the template PPT file.
    Dim targetPresentation As Object
    Set targetPresentation = GetOrOpenPresentation(templatePPTFullName)
    If targetPresentation Is Nothing Then
        Debug.Print "Template file `" & templatePPTFullName & "` was not found."
        MsgBox "Template file `" & templatePPTFullName & "` was not found."
        GoTo Finally
    End If
    
    Dim saveFullName As String
    saveFullName = GetSaveFullName(ThisWorkbook)
    
    ' Saves the template PPT file as a new file.
    ' The new file will be input into `targetPresentation` variable.
    targetPresentation.SaveAs saveFullName
    
    ' Gets the number of slides in the template PPT file.
    '`slidesCount` variable has the number of slides at this moment.
    Dim slidesCount As Long
    slidesCount = targetPresentation.Slides.Count
    
    Dim ST As Double: ST = Timer
    
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
    
    PPT.Visible = True
    
    Debug.Print Timer - ST, "end of begin"
    
    
    ''''''''''''''''''''''''''''''
    'process
    
    Application.StatusBar = "0 / " & UBound(quizList, 1) - 1
    
    Dim i As Long
    For i = 2 To UBound(quizList, 1)
        
        Dim ST2 As Double: ST2 = Timer
        
        If CLng(quizList(i, 1)) > slidesCount Then
            GoTo Continue
        End If
        
        Dim templateSlide As Object
        Set templateSlide = targetPresentation.Slides(CLng(quizList(i, 1)))
        
        Dim targetSlide As Object
        Set targetSlide = templateSlide.Duplicate
        targetSlide.MoveTo targetPresentation.Slides.Count
        targetSlide.SlideShowTransition.Hidden = msoFalse
        
        Call ClearEffects(targetSlide)
        Call CopyEffects(templateSlide, targetSlide, CLng(quizList(i, 2)))
        
        Dim titleShape As Object
        Set titleShape = GetShapeByText(targetSlide, "{title}")
        If Not titleShape Is Nothing Then
            titleShape.TextFrame.TextRange.Text = quizList(i, 3)
            Set titleShape = Nothing
        End If
        
        Dim textShape As Object
        Set textShape = GetShapeByText(targetSlide, "{Q}")
        If Not textShape Is Nothing Then
            textShape.TextFrame.TextRange.Text = quizList(i, 4)
            Set textShape = Nothing
        End If
        
        Dim j As Long
        For j = 5 To UBound(quizList, 2)
            
            Dim choiceIndex As Long
            choiceIndex = j - 4
            
            Dim choiceShape As Object
            Set choiceShape = GetShapeByText(targetSlide, "{" & choiceIndex & "}")
            If Not choiceShape Is Nothing Then
                choiceShape.TextFrame.TextRange.Text = quizList(i, j)
                Set choiceShape = Nothing
            End If
            
        Next j
        
        Debug.Print Timer - ST2, "loop " & i - 1
        
Continue:
        Application.StatusBar = i - 1 & " of " & UBound(quizList, 1) - 1
        
    Next i
    
    Debug.Print Timer - ST, "end of process"
    
    
    '''''''''''''''''''''''''''
    'end
    
    targetPresentation.Save
    
    Debug.Print Timer - ST, "end"
    Debug.Print
    
    ThisWorkbook.Activate
    MsgBox "Finished."
    
Finally:
    If Err.Number > 0 Then MsgBox Err.Number & vbCrLf & Err.Description
    
    Application.StatusBar = False
    
End Sub


'Private Function GetSelectedPresentation() As Object
'
'    Dim targetFullName As String
'    targetFullName = Application.GetOpenFilename("PowerPoint Presentation,*.pptx,PowerPoint 97-2003 Presentaion,*.ppt")
'
'    If targetFullName = "False" Then
'        Exit Function
'    End If
'
'    Dim PPT As Object
'    Set PPT = CreateObject("PowerPoint.Application")
'
'    Dim PR As Object
'    For Each PR In PPT.Presentations
'        If PR.FullName = targetFullName Then
'            Set GetSelectedPresentation = PR
'            Set PPT = Nothing
'            Exit Function
'        End If
'    Next PR
'
'    Set GetSelectedPresentation = PPT.Presentations.Open(targetFullName, ReadOnly:=True)
'    Set PPT = Nothing
'
'End Function


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


Private Function GetSaveFullName(quiz_list_workbook As Workbook) As String
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Do
        Dim result As String
        result = quiz_list_workbook.Path & "\" & FSO.GetBaseName(quiz_list_workbook.FullName) & Format(Now, "_yyyy-MM-dd_HH-mm-ss") & ".pptx"
    Loop While Dir(result) <> ""
    
    Set FSO = Nothing
    
    GetSaveFullName = result
    
End Function


Private Function GetShapeByText(target_slide As Object, target_text As String) As Object
    
    Dim SHP As Object
    For Each SHP In target_slide.Shapes
        
        If SHP.TextFrame.TextRange.Text = target_text Then
            Set GetShapeByText = SHP
            Exit Function
        End If
        
    Next SHP
    
End Function


Private Sub ClearEffects(target_slide As Object)
    
    With target_slide.TimeLine.MainSequence
        Dim i As Long
        For i = .Count To 1 Step -1
            .Item(i).Delete
        Next i
    End With
    
End Sub


Private Sub CopyEffects(template_slide As Object, target_slide As Object, correct_choice_index As Long)
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\{(\d+)\}"
    
    Dim EF As Object
    For Each EF In template_slide.TimeLine.MainSequence
        
        Dim shapeText As String
        shapeText = EF.Shape.TextFrame.TextRange.Text
        
        ' matching the shape text to judge the text has a number
        Dim matchResult As Object
        Set matchResult = regex.Execute(shapeText)
        
        Dim textForSearchingShape As String
        
        If matchResult.Count = 0 Then
            ' didn't match, meaning the shape had no number
            textForSearchingShape = shapeText
        Else
            ' matched, meaning the shape had a number
            Dim numberInBraces As Long
            numberInBraces = CLng(matchResult(0).SubMatches(0))
            
            ' shifting the shape number
            If numberInBraces = 1 Then
                textForSearchingShape = "{" & correct_choice_index & "}"
            Else
                If numberInBraces <= correct_choice_index Then
                    textForSearchingShape = "{" & numberInBraces - 1 & "}"
                End If
            End If
            
        End If
        
        ' getting the effect to add effect to
        Dim shapeToAddEffectTo As Object
        Set shapeToAddEffectTo = GetShapeByText(target_slide, textForSearchingShape)
        
        ' adding a new effect to the shape
        Dim newEf As Object
        Set newEf = target_slide.TimeLine.MainSequence.AddEffect(shapeToAddEffectTo, EF.EffectType, , EF.Timing.TriggerType)
        
        With newEf
        
'            .Behaviors.Item(1).Type = ef.Behaviors.Item(1).Type
            
'            .EffectInformation.AfterEffect = ef.EffectInformation.AfterEffect
'            .EffectInformation.TextUnitEffect = ef.EffectInformation.TextUnitEffect
            
            If EF.Exit = -1 Then
                .Exit = EF.Exit
            End If
            
            .Timing.Accelerate = EF.Timing.Accelerate
            .Timing.AutoReverse = EF.Timing.AutoReverse
            .Timing.BounceEnd = EF.Timing.BounceEnd
            .Timing.BounceEndIntensity = EF.Timing.BounceEndIntensity
            .Timing.Decelerate = EF.Timing.Decelerate
            .Timing.Duration = EF.Timing.Duration
            .Timing.RepeatCount = EF.Timing.RepeatCount
            .Timing.RepeatDuration = EF.Timing.RepeatDuration
            .Timing.Restart = EF.Timing.Restart
            .Timing.RewindAtEnd = EF.Timing.RewindAtEnd
            .Timing.SmoothEnd = EF.Timing.SmoothEnd
            .Timing.SmoothStart = EF.Timing.SmoothStart
            .Timing.Speed = EF.Timing.Speed
'            .Timing.TriggerBookmark = ef.Timing.TriggerBookmark
            .Timing.TriggerDelayTime = EF.Timing.TriggerDelayTime
            
        End With
        
    Next EF
    
End Sub


Public Sub AddToContextMenu()
    
    With Application.CommandBars
        
        Dim i As Long
        For i = 1 To .Count
            
            If .Item(i).Name = "Cell" Then
                
                With .Item(i).Controls.Add(Type:=msoControlPopup, Temporary:=True)
                    .BeginGroup = True
                    .Caption = "&" & ThisWorkbook.Name

                    With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                        .Caption = "Make &Quiz"
                        .OnAction = ThisWorkbook.Name & "!MakeQuiz"
                    End With
                    
                End With
                
            End If
            
        Next i
        
    End With
    
End Sub
