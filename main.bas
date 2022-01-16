Attribute VB_Name = "main"
Option Explicit

Public Sub MakeQuiz()
    
    ' begin
    
    On Error Resume Next
    
    Dim ST As Double: ST = Timer
    
    Dim saveFullName As String
    saveFullName = GetSaveFullName(ActiveWorkbook)
    
    Dim PPT As Object
    Set PPT = CreateObject("PowerPoint.Application")
    
    Dim targetPresentation As Object
    Set targetPresentation = GetSelectedPresentation
    If targetPresentation Is Nothing Then
        MsgBox "Suspended"
        Exit Sub
    End If
    
    targetPresentation.SaveAs saveFullName
    
    Application.StatusBar = "starting..."
    
    PPT.Visible = True
    
    Debug.Print Timer - ST, "end of begin"
    
    
    ''''''''''''''''''''''''''''''
    'process
    
    Dim quizList As Variant
    quizList = ActiveWorkbook.ActiveSheet.Cells(1, 1).CurrentRegion.Value
    
    Dim slidesCount As Long
    slidesCount = targetPresentation.Slides.Count
    
    Application.StatusBar = "0 / " & UBound(quizList, 1) - 1
    
    Dim i As Long
    For i = 2 To UBound(quizList, 1)
        
        Dim ST2 As Double: ST2 = Timer
        
        If CLng(quizList(i, 1)) > slidesCount Then
            '
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
        
        Application.StatusBar = i - 1 & " of " & UBound(quizList, 1) - 1
        
    Next i
    
    Debug.Print Timer - ST, "end of process"
    
    
    '''''''''''''''''''''''''''
    'end
    
    targetPresentation.Save
    
    Debug.Print Timer - ST, "end"
    Debug.Print
    
    ActiveWorkbook.Activate
    MsgBox "Finished."
    
Finally:
    If Err.Number > 0 Then MsgBox Err.Number & vbCrLf & Err.Description
    
    Application.StatusBar = False
    
End Sub


Private Function GetSelectedPresentation() As Object
    
    Dim targetFullName As String
    targetFullName = Application.GetOpenFilename("PowerPoint Presentation,*.pptx,PowerPoint 97-2003 Presentaion,*.ppt")
    
    If targetFullName = "False" Then
        Exit Function
    End If
    
    Dim PPT As Object
    Set PPT = CreateObject("PowerPoint.Application")
    
    Dim pr As Object
    For Each pr In PPT.Presentations
        If pr.FullName = targetFullName Then
            Set GetSelectedPresentation = pr
            Set PPT = Nothing
            Exit Function
        End If
    Next pr
    
    Set GetSelectedPresentation = PPT.Presentations.Open(targetFullName, ReadOnly:=True)
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
    
    Dim EF As Object
    For Each EF In template_slide.TimeLine.MainSequence
        
        Dim targetEffectText As String
        targetEffectText = EF.Shape.TextFrame.TextRange.Text
        
        Dim targetEffectTextWithoutBraces As String
        targetEffectTextWithoutBraces = Replace(targetEffectText, "{", "")
        targetEffectTextWithoutBraces = Replace(targetEffectTextWithoutBraces, "}", "")
        
        If IsNumeric(targetEffectTextWithoutBraces) Then
            
            If targetEffectTextWithoutBraces = 1 Then
                targetEffectText = "{" & correct_choice_index & "}"
            Else
                If targetEffectTextWithoutBraces <= correct_choice_index Then
                    targetEffectText = "{" & CLng(targetEffectTextWithoutBraces) - 1 & "}"
                End If
            End If
            
        End If
        
        Dim targetShape As Object
        Set targetShape = GetShapeByText(target_slide, targetEffectText)
        
        Dim tmpEffectType As Long
        tmpEffectType = EF.EffectType
        
        Dim newEf As Object
        Set newEf = target_slide.TimeLine.MainSequence.AddEffect(targetShape, EF.EffectType, , EF.Timing.TriggerType)
        
        Dim tmpBehaviorsType As Long
        tmpBehaviorsType = EF.Behaviors.Item(1).Type
        
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