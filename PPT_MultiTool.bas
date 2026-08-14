Attribute VB_Name = "PPT_MultiTool_v1_1"
Option Explicit
'**
' Version 1.1
' License information:
' https://github.com/c-s-code/powerpoint-fixer/blob/main/LICENSE
'**

Private Type SlideInfo
    SlideNo As Integer
    time As Double
End Type

Private Type IssueInfo
    count As Integer
    time As Double
    animatedTransitions As Integer
    hiddenSlides As Integer
    slideFooters As Integer
    timingOff As Integer
    multiAnim As Collection
    zeroAnim As Collection
End Type
    
Private Type PresentationInfo
    runTime As Double
    slideCount As Integer
End Type

Sub SlideChecker()
    '**
    'moves all audio to bottom right (ish), sets audio volume to 100% and sets the image as decorative
    'NB about COLLECTIONS (lists) - list indices start at 1
    '**
    Dim pres as Presentation
    Set pres = ActivePresentation

    Debug.Print (vbNewLine & "======================" & vbNewLine)
    Debug.Print ("Processing '" & pres.Name & "'")
    
    Dim slide As slide
    Dim shape As shape

    Dim presInfo As PresentationInfo
    With presInfo:
        .runTime = 0
        .slideCount = 0
    End With

    'used on first image found, run compress all images, then turn off feature
    Dim runOnceToCompress As Boolean
    runOnceToCompress = False 'feature in testing, broken

    'issue tracking
    Dim issueTracker As IssueInfo
    With issueTracker:
        .count = 0
        .time = 0
        .animatedTransitions = 0
        .hiddenSlides = 0
        .slideFooters = 0
        .timingOff = 0
        Set .multiAnim = New Collection
        Set .zeroAnim = New Collection
    End With
    
    Dim longSlide As SlideInfo
    Dim shortSlide As SlideInfo
    longSlide.time = 0
    shortSlide.time = 1000000
    
    Dim maxTransitionTolerance As Double
    maxTransitionTolerance = 0.5
    
    Dim narrationSlowMode as Boolean
    narrationSlowMode = (MsgBox("Slow mode is for lectures recorded correctly, but slowly." & vbNewLine & "(With longer pauses between slides - PPT seems to trim a little off automagically)" & vbNewLine & vbNewLine & "Choose 'Yes' to increase the automatic transition tolerance by 1.5 seconds." & vbNewLine & "Choose 'No' to keep default 0.5s.", vbYesNo, "Enable Slow Mode?") = vbYes)
    If narrationSlowMode Then
        maxTransitionTolerance = maxTransitionTolerance + 1.5
    End If
    
    'look ahead for movies, disallow auto-fix if true
    Dim autoFixAll As Boolean
    Dim moveAudioOff As Boolean
    
    Dim isAudioNarration As Boolean
    
    Dim hasMovies as Boolean
    hasMovies = HasMovieLookAhead(pres)
    Dim hasAudio as Boolean
    hasAudio = HasAudioLookAhead(pres)

    'ask if is video or audio narration
    isAudioNarration = (MsgBox("Does this presentation use audio narration?" & vbNewLine & vbNewLine & "Choose 'No' to indicate video narration.", vbYesNo, "Narration Type") = vbYes)
    
    Dim continueScript As Boolean
    continueScript = True
    
    If isAudioNarration Then
        If hasMovies Then
            autoFixAll = False
            continueScript = (MsgBox("Found a movie in the pre-check. See 'immediate' window for more info." & vbNewLine & "Continue script?", vbYesNo, "Continue?") = vbYes)
            If Not continueScript Then
                End
            End If
        Else
            autoFixAll = (MsgBox("Would you like to auto-fix all issues." & vbNewLine & vbNewLine & "This may result in unforseen issues. You have been warned!", vbYesNo, "Autofix?") = vbYes)
        End If
    Else
        If hasAudio Then
            autoFixAll = False
            continueScript = (MsgBox("Found audio file in the pre-check. See 'immediate' window for more info." & vbNewLine & "Continue script?", vbYesNo, "Continue?") = vbYes)
            If Not continueScript Then
                End
            End If
        Else
            autoFixAll = (MsgBox("Would you like to auto-fix all issues." & vbNewLine & vbNewLine & "This may result in unforseen issues. You have been warned!", vbYesNo, "Autofix?") = vbYes)
        End If
    End If
    
    moveAudioOff = (MsgBox("Move all media off screen?" & vbNewLine & vbNewLine & "Choose 'No' to move audio to bottom right, and leave video where it is.", vbYesNo, "Audio Location") = vbYes)
       

    'loop through each slide in presentation
    For Each slide In pres.Slides
        
        'make slide the active window
        ActiveWindow.View.GotoSlide slide.SlideIndex
    
        'use collections (list) to store all valid shapes, must clear w/nothing otherwise it accumulates even when using "Dim var = New Collection" inside loop
        Dim audioList As Collection
        Dim videoList As Collection 'create
        Set audioList = Nothing
        Set videoList = Nothing 'clear
        Set audioList = New Collection
        Set videoList = New Collection 'initialise
        
        Dim moveit As Boolean
        Dim isNarration As Boolean
        Dim isDelete As Boolean
        
        Dim slideTime As Double
        slideTime = 0
        
        '!= 1 animation slide check
        Select Case CountSlideAnimations(slide)
        Case 0
            issueTracker.count = issueTracker.count + 1
            issueTracker.zeroAnim.Add (slide.SlideIndex)
            Debug.Print ("added an issue CountSlideAnimations case 0")
        Case Is > 1
            issueTracker.count = issueTracker.count + 1
            issueTracker.multiAnim.Add (slide.SlideIndex)
            Debug.Print ("added an issue CountSlideAnimations case > 1")
        Case Else
            'do nothing
        End Select

        'Remove slide day/date/footer/slideno if they are visible on slides (will not remove manual stuff)
        If RemoveSlideFooters(slide) Then
            issueTracker.slideFooters = issueTracker.slideFooters + 1
            issueTracker.count = issueTracker.count + 1
            Debug.Print ("added an issue RemoveSlideFooters")
        End If
            
        'loop through each shape (object) in slide
        For Each shape In slide.Shapes
        
            'should run once on first image found and apply to all images
            If runOnceToCompress Then
                If shape.Type = msoPicture Then
                    runOnceToCompress = CompressImages()
                End If
            End If

            'check if its audio or video
            If IsMedia(shape) Then
                audioList.Add shape
            ElseIf IsMedia(shape, True) Then
                videoList.Add shape
            End If
               
        Next shape
        
        'logic to check what to do now
        If (audioList.count + videoList.count) = 1 Then
        'case for single piece of media
            If audioList.count = 1 Then
                'audio
                MoveMedia pres, audioList(1), autoFixAll, moveAudioOff
                SetMaxVolume audioList(1)
                FixAnimationSettings audioList(1), slide
                slideTime = UpdateSlideTime(audioList(1), slideTime)
            Else
                'video
                MoveMedia pres, videoList(1), autoFixAll, moveAudioOff, True, Not isAudioNarration
                SetMaxVolume videoList(1)
                FixAnimationSettings videoList(1), slide
                slideTime = UpdateSlideTime(videoList(1), slideTime)
            End If
        
        ElseIf (audioList.count + videoList.count) > 1 Then
        'multiple media on page - warning for manual check

            MsgBox "Multiple media files on page - please review!", vbInformation, "Info"
                
            'audio foreach
            Dim audioShape as Variant
            For Each audioShape In audioList
                
                Dim ashape As shape
                Set ashape = audioShape 'VBA doesnt do implicit type inheritance?
                ashape.Select
                
                isNarration = YesNoDialog(False, "Object: " & ashape.Name & vbNewLine & "Is this audio narration?", "Query")
            
                If isNarration = True Then
                    slideTime = UpdateSlideTime(ashape, slideTime)
                    SetMaxVolume ashape
                    FixAnimationSettings ashape, slide
                Else
                'if not narration, ask about removing anim
                    isDelete = YesNoDialog(False, "Remove animation from this non-narration object?", "Careful!!")
                    If isDelete Then
                        ashape.AnimationSettings.Animate = msoFalse
                        ashape.AnimationSettings.AdvanceMode = msoFalse
                    End If
    
                End If
                
                moveit = YesNoDialog(autoFixAll, "Object: " & ashape.Name & vbNewLine & vbNewLine & "Move this audio shape to bottom right?" & vbNewLine & "Select NO to leave in place", "Move Object?")
                
                If moveit = True Then
                    MoveMedia pres, ashape, autoFixAll, moveAudioOff
                Else
    
                'do nothing
                End If
               
            Next audioShape
            
            Dim videoShape as Variant
            For Each videoShape In videoList
                
                Dim vshape As shape
                Set vshape = videoShape 'VBA doesnt do implicit type inheritance?
                vshape.Select
                
                'is video narration?
                isNarration = YesNoDialog(False, "Object: " & vshape.Name & vbNewLine & "Is this video narration?" & vbNewLine & "Select NO if video is played under/alongside audio narration", "Query")
                'isNarration = MsgBox("Object: " & vshape.Name & vbNewLine & "Is this video narration?", vbYesNo, "Query")
                
                If isNarration = True Then
                    slideTime = UpdateSlideTime(vshape, slideTime)
                    MoveMedia pres, vshape, False, moveAudioOff, True, Not isAudioNarration
                Else
                'if not narration, ask about removing anim
                    isDelete = YesNoDialog(False, "Remove animation from this non-narration object?" & vbNewLine & "Select NO to retain autoplay or on-click play settings", "Careful!!")
                    If isDelete Then
                        vshape.AnimationSettings.Animate = msoFalse
                    End If
                End If
            Next videoShape

        Else
        'no media on page - remove transition timings etc
            slideTime = 5 * 1000
        End If

        'check and fix slide timings with respect to audio length                     
        Dim tempTimingValue As IssueInfo
        tempTimingValue = CheckTimings(slideTime, slide, autoFixAll, maxTransitionTolerance)
        
        'update issue info tracker
        With issueTracker:
            .count = issueTracker.count + tempTimingValue.count
            Debug.Print ("added " & tempTimingValue.count & " count of issues to tracker")
            .time = issueTracker.time + tempTimingValue.time
            .timingOff = issueTracker.timingOff + tempTimingValue.timingOff
            .animatedTransitions = issueTracker.animatedTransitions + RemoveAnimatedTransitions(slide)
        End With
        
        'update runtime only if not hidden
        If Not slide.SlideShowTransition.Hidden Then
            presInfo.slideCount = presInfo.slideCount + 1
            presInfo.runTime = presInfo.runTime + slideTime
            'runTime = runTime + slideTime
            
            'a slide can be both longest and shortest if it is the only slide, so no if-else
            If slideTime > longSlide.time Then
                With longSlide:
                .SlideNo = slide.SlideIndex
                .time = slideTime
                'Debug.Print ("updated longest slide on slide no: " & slide.SlideIndex & " - with value: " & slideTime)
                End With
            End If
            If slideTime < shortSlide.time Then
                With shortSlide:
                .SlideNo = slide.SlideIndex
                .time = slideTime
                End With
            End If
            
        Else
            issueTracker.hiddenSlides = issueTracker.hiddenSlides + 1
            Debug.Print ("slide " & slide.SlideIndex & " is hidden in slideshow.")
        End If
    
    Next slide
    
    ShowSummary presInfo.runTime, longSlide, shortSlide, presInfo.slideCount, issueTracker ', multiAnimationSlides, zeroAnimationSlides
    
    Debug.Print ("issue tracker data: " & issueTracker.count & issueTracker.hiddenSlides & issueTracker.multiAnim.count & issueTracker.zeroAnim.count & issueTracker.animatedTransitions & issueTracker.timingOff)
    'ask to export
    ExportFiles pres

    MsgBox "Processing completed.", vbInformation, "Done"
End Sub


'**
'subroutines go below here
'**


Sub ExportFiles(pres as Presentation)
'ask to export mp4 - must export PDF manually to preserve accessible header tags etc

    If YesNoDialog(False, "Do you want to export this presentation to video?", "Export files?") Then
    'get save location
        Dim folderPath As String
        Dim saveName As String
        'Dim targetPath As String
        
        'strip off .pptx from file name
        saveName = Left(pres.Name, InStrRev(pres.Name, ".") - 1)
        folderPath = GetFolder()
        
        If folderPath = "" Then
            MsgBox "You didn't select a save location. Export cancelled.", vbInformation, "Process Cancelled"
            Exit Sub
        End If

        ExportVid pres, folderPath, saveName
        
    Else
        Debug.Print ("Nothing exported.")
    End If
End Sub


Sub ShowSummary(time As Double, longest As SlideInfo, shortest As SlideInfo, numSlides As Integer, tracker As IssueInfo)
    'display the run time in minutes and seconds
    Dim info As String
    Dim furtherInfo As String
    
    info = "Total expected run time: " & GetTimeMinSec(time)
    furtherInfo = vbNewLine & vbNewLine & "Number of slides: " & numSlides & vbNewLine & "Longest slide: " & GetTimeMinSec(longest.time) & " (slide " & longest.SlideNo & ")"
    
    If tracker.hiddenSlides > 0 Then
            furtherInfo = furtherInfo & vbNewLine & "Number of hidden slides: " & tracker.hiddenSlides & " (not counted in timings)"
    End If
        
    furtherInfo = furtherInfo & vbNewLine & "Shortest slide: " & GetTimeMinSec(shortest.time) & " (slide " & shortest.SlideNo & ")"
    furtherInfo = furtherInfo & vbNewLine & "Avg slide length: " & GetTimeMinSec(time / numSlides)
    furtherInfo = furtherInfo & vbNewLine & vbNewLine & "Number of issues found: " & tracker.count + tracker.animatedTransitions
    furtherInfo = furtherInfo & vbNewLine & "... of which..."
    furtherInfo = furtherInfo & vbNewLine & " - animated transitions: " & tracker.animatedTransitions
    furtherInfo = furtherInfo & vbNewLine & " - day/date/footer/slideno visible: " & tracker.slideFooters
    furtherInfo = furtherInfo & vbNewLine & " - hidden slides (check): " & tracker.hiddenSlides
    
    If tracker.zeroAnim.count > 0 Then
        furtherInfo = furtherInfo & vbNewLine & " - zero animations (check): " & tracker.zeroAnim.count & " sllides: " & MakeCollectionString(tracker.zeroAnim)
    End If
    
    If tracker.multiAnim.count > 0 Then
        furtherInfo = furtherInfo & vbNewLine & " - multiple animations (check): " & tracker.multiAnim.count & " sllides: " & MakeCollectionString(tracker.multiAnim)
    End If
    
    furtherInfo = furtherInfo & vbNewLine & "Total slide transition discrepancy: " & GetTimeMinSec(tracker.time * 1000)
    
    MsgBox info & furtherInfo, vbInformation, "Presentation Summary"
    Debug.Print (info & furtherInfo)
    'Debug.Print (tracker.animatedTransitions & tracker.count & tracker.hiddenSlides & tracker.slideFooters & tracker.time)
End Sub


Sub MoveMedia(pres As Presentation, shape As shape, autoFixAll As Boolean, audioOffScreen As Boolean, Optional isMovie As Boolean = False, Optional movieNarration As Boolean = False)
    '**
    'move media to the position required
    'powerpoint uses a 1/2 pixel scale, so 1920x1080 = 960x540 "powerpoint pixels" slide size
    'shape positions (and ppt positions) are relative to top left (0,0) of screen & shape
    '**
    On Error GoTo ErrorHandler

    'offset fudge - assume image plus buffer is approx 64 by 64
    Dim h As Integer
    Dim w As Integer
    Dim h_offset As Integer
    Dim w_offset As Integer
    Dim moveOffPage As Boolean
    Dim manual_tweak_h as Single
    Dim manual_tweak_w as Single
    
    h = shape.Height
    w = shape.Width
    'div by two because we are offsetting centre of obj
    h_offset = Int((64 - h) / 2)
    w_offset = Int((70 - w) / 2) 'a little extra buffer for width
    
    'change these two values to add a manual offset, (default = 0)
    manual_tweak_h = 30
    manual_tweak_w = 5
    
    If isMovie = False Then
        If audioOffScreen = True Then
            shape.Left = pres.PageSetup.SlideWidth + 100
        Else
            shape.Top = pres.PageSetup.SlideHeight - h - h_offset - manual_tweak_h
            shape.Left = pres.PageSetup.SlideWidth - w - w_offset - manual_tweak_w
        End If
    Else
        If movieNarration Then
            If audioOffScreen Then
                shape.Left = pres.PageSetup.SlideWidth + 100
            End If
        Else
            moveOffPage = YesNoDialog(autoFixAll, "Object: " & shape.Name & vbNewLine & "Move video off the page?", "Move Object?")
            If moveOffPage Then
                shape.Left = pres.PageSetup.SlideWidth + 100
            End If
        End If
    End If

    Exit Sub

ErrorHandler:
    ErrorMsg Err.Description, "MoveMedia"
End Sub


Sub SetMaxVolume(shape As shape)
    'set audio vol to max
    
    Dim ashape As shape
    Set ashape = shape
    
    On Error GoTo ErrorHandler

    ashape.MediaFormat.Volume = 1
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err.Description, "SetMaxVolume"
End Sub


Sub FixAnimationSettings(shape As shape, slide As slide)

    On Error GoTo ErrorHandler

    shape.AnimationSettings.PlaySettings.PlayOnEntry = True
    shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = True

    Exit Sub

ErrorHandler:
    ErrorMsg Err.Description, "FixAnimationSettings"   
End Sub

Function CountSlideAnimations(slide As slide) As Integer
    CountSlideAnimations = 1
    
    On Error GoTo ErrorHandler

    If slide.TimeLine.MainSequence.count > 1 Then
        CountSlideAnimations = 2
    ElseIf slide.TimeLine.MainSequence.count = 0 Then
        CountSlideAnimations = 0
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorMsg Err.Description, "CountSlideAnimations"
End Function

Function RemoveAnimatedTransitions(slide As slide) As Integer
    'remove any daft animated transitions
    
    RemoveAnimatedTransitions = 0
    On Error GoTo ErrorHandler
    If Not slide.SlideShowTransition.EntryEffect = ppEffectNone Then
        slide.SlideShowTransition.EntryEffect = ppEffectNone
        RemoveAnimatedTransitions = 1
    End If
    
    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "RemoveAnimatedTransitions"

End Function

Function MakeImageDecorative(shape As shape, slide As slide) As Integer
    'fix decorative settings on non-text elements
    'not specifically searching for msoPicture etc to allow for exceptions
    
    MakeImageDecorative = 0
    On Error GoTo ErrorHandler
        
    If shape.HasTextFrame = msoTrue Then
    
        If shape.TextFrame.TextRange.Text = "" Then
            Debug.Print (shape.Name & " contains no text.")
            If shape.AlternativeText = "" Then
                
                shape.Decorative = msoTrue
                MakeImageDecorative = 1
                Debug.Print ("Shape has no alt text. Made " & shape.Name & " decorative")
                Exit Function
            Else
                shape.Decorative = msoFalse
                Debug.Print ("Shape contains alt text. Made " & shape.Name & " NOT decorative")
            End If
        Else
            shape.Decorative = msoFalse
            Debug.Print ("Shape contains text. Made " & shape.Name & " NOT decorative")
        End If
        
    Else
        If shape.AlternativeText = "" Then
            shape.Decorative = msoTrue
            MakeImageDecorative = 1
            Exit Function
        Else
            shape.Decorative = msoFalse
        End If
        
    End If
    
    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "MakeImageDecorative"
End Function

Function CheckTimings(audioTime As Double, slide As slide, skipMessage As Boolean, tolerance As Double) As IssueInfo
    'compare transition time - warning if mismatch
    CheckTimings.count = 0
    CheckTimings.time = 0
    CheckTimings.timingOff = 0
    
    On Error GoTo ErrorHandler
    
    'Dim length As Double
    Dim slideDuration As Double
    Dim transitionDiff As Double
    Dim updateTransition As Boolean
    
    'get transition time
    With slide.SlideShowTransition
        slideDuration = .AdvanceTime
        'set advance on time checkbox to true
        .AdvanceOnTime = msoTrue
    End With
    'warn if audio-transition difference is over 1 sec
    transitionDiff = Abs(audioTime / 1000 - slideDuration)
    If transitionDiff > tolerance Then
        Dim warningText As String
        Dim dialogTitle As String
        
        warningText = "Transition: " & Round(slideDuration, 2) & "  Audio Time: " & Round((audioTime / 1000), 2) & vbNewLine
        warningText = warningText & "Transition and narration timings: " & Round(transitionDiff, 2) & " second difference detected." & vbNewLine & "Do you want to update transition time to match narration?"
        dialogTitle = "Slide index " & slide.SlideIndex & ": Update Timings?"
        
        'ask if updating transition time
        updateTransition = YesNoDialog(skipMessage, warningText, dialogTitle)
        If updateTransition = True Then
            CheckTimings.count = 1
            CheckTimings.time = transitionDiff
            CheckTimings.timingOff = 1
            slide.SlideShowTransition.AdvanceTime = CSng(Round(audioTime / 1000, 2))
        End If
        
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorMsg Err.Description, "CheckTimings"

End Function


'**
'functions below here
'**

Function YesNoDialog(autofix As Boolean, prompt As String, Optional windowTitle = "Information") As Boolean
    'create yes no dialog window to confirm action unless autofix is chosen
    If autofix = True Then
        YesNoDialog = True
        Exit Function
    Else
        YesNoDialog = (MsgBox(prompt, vbYesNo, windowTitle) = vbYes)
        Exit Function
    End If
End Function

Function UpdateSlideTime(shape As shape, currentSlideTime As Double) As Double
    'update the slide time total
    UpdateSlideTime = currentSlideTime

    On Error GoTo ErrorHandler

    'update run time for media that is trimmed
    Dim mediaStart As Double
    Dim mediaEnd As Double
    mediaStart = shape.MediaFormat.StartPoint
    mediaEnd = shape.MediaFormat.EndPoint
    UpdateSlideTime = mediaEnd - mediaStart + currentSlideTime
    
    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "UpdateSlideTime"   
End Function

Function IsMedia(shape As shape, Optional checkMovie As Boolean = False) As Boolean
    'return true if shape is a media file
    IsMedia = False

    On Error GoTo ErrorHandler

    If shape.Type = msoMedia Then
        If checkMovie = False Then
            If shape.MediaType = ppMediaTypeSound Then
                IsMedia = True
                'Debug.Print ("Found a sound file!")
                Exit Function
            End If
        Else
            If shape.MediaType = ppMediaTypeMovie Then
                IsMedia = True
                'Debug.Print ("Found a video file!")
                Exit Function
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "IsMedia"       
End Function

Function HasMovieLookAhead(pres as Presentation) As Boolean
    On Error GoTo ErrorHandler    

    Dim slide As slide
    Dim shape As shape
    HasMovieLookAhead = False

    For Each slide In pres.Slides
        'loop through each shape (object) in slide
        For Each shape In slide.Shapes
            'check if it has video
            If IsMedia(shape, True) Then
                HasMovieLookAhead = True
                Debug.Print ("Found movie on slide " & slide.SlideIndex & ", aborting movie look-ahead.")
                Exit Function
            End If
        Next shape
    Next slide

    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "HasMovieLookAhead"      
End Function

Function HasAudioLookAhead(pres as Presentation) As Boolean
    On Error GoTo ErrorHandler

    Dim slide As slide
    Dim shape As shape
    HasAudioLookAhead = False
    For Each slide In pres.Slides
        For Each shape In slide.Shapes
            'check if it has audio
            If IsMedia(shape) Then
                HasAudioLookAhead = True
                Debug.Print ("Found audio on slide " & slide.SlideIndex & ", aborting audio look-ahead.")
                Exit Function
            End If
        Next shape
    Next slide

    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "HasAudioLookAhead"      
End Function

Function GetTimeMinSec(time As Double) As String
    On Error GoTo ErrorHandler

    Dim mins As Integer
    mins = Int(time / 60000)
    Dim sec As Integer
    sec = Int(time / 1000) Mod 60
    
    If mins = 0 Then
        GetTimeMinSec = sec & "s"
    Else
        GetTimeMinSec = mins & "m " & sec & "s"
    End If

    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "GetTimeMinSec"      
End Function


''****
'media compression
'****
Function CompressImages() As Boolean
'sources:
' http://www.vbaexpress.com/forum/showthread.php?63592-How-to-start-Compress-Pictures-from-VBA
' https://stackoverflow.com/questions/76542615/how-to-compress-pictures-as-i-paste-them-word-for-windows
    CompressImages = True

    On Error GoTo ErrorHandler

    'some issue here but no idea about solution
    Application.CommandBars.ExecuteMso ("PicturesCompress")
    
    SendKeys "%a", True
    SendKeys "%w", True
    SendKeys "{ENTER}", True

    CompressImages = False

Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "CompressImages"
End Function


'****
'media export subroutines and functions below here
'****
Sub ExportVid(ppPres As Presentation, path As String, FileName As String)
'export mp4 to target location
    On Error GoTo ErrorHandler
    Dim openPres As PowerPoint.Presentation
    Dim targetPath as String

    'check every instance of ppt that is running for exports - they can actually be queued!
    For Each openPres In PowerPoint.Presentations
        If openPres.CreateVideoStatus = ppMediaTaskStatusInProgress Then
            MsgBox "There is already a video export in progress.", vbCritical, "Warning - multiple exports attempted!"
            Debug.Print ("Aborted video export.")
            Exit Sub
        End If
    Next openPres
    
    targetPath = path & GetOSDivider() & FileName & ".mp4"
    
    If Not ppPres Is Nothing Then
        Debug.Print ("Creating the following video: " & targetPath)
        ppPres.CreateVideo FileName:=targetPath, UseTimingsAndNarrations:=True, VertResolution:=720, FramesPerSecond:=30, Quality:=100
    Else
        MsgBox "Could not start export.", vbCritical, "Error!"
        Debug.Print ("Could not create video.")
    End If
    
Exit Sub
    
ErrorHandler:
    ErrorMsg Err.Description, "ExportVid"
End Sub


Function GetFolder() As String
'user selects folder to save to
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = ""
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Function GetOSDivider() As String
'get correct OS path divider backslash windows, forward mac
    Dim os As String
    os = Application.OperatingSystem
    If InStr(1, os, "windows", vbTextCompare) Then
        GetOSDivider = "\"
    Else
        GetOSDivider = "/"
    End If
End Function

Function RemoveSlideFooters(sld As slide) As Boolean
    RemoveSlideFooters = False
    On Error GoTo ErrorHandler
    
    'Debug.Print ("Checking for slide numbers in slide " & slide.SlideIndex & ".")
    
    With sld.HeadersFooters
        If .DateAndTime.Visible = msoTrue Or .Footer.Visible = msoTrue Or .SlideNumber.Visible = msoTrue Then
            'Debug.Print ("Found footer info day/date/footer/slideno! Removing...")
            RemoveSlideFooters = True
        End If
    End With
    
    sld.HeadersFooters.Clear 'actually clears the footers
    
    Exit Function
    
ErrorHandler:
    ErrorMsg Err.Description, "RemoveSlideFooters"
End Function

Function MakeCollectionString(col As Collection) As String
    MakeCollectionString = ""
    On Error GoTo ErrorHandler

    Dim myList As String
    Dim first As Boolean
    Dim Item as Variant

    first = True
    For Each Item In col
        If first Then
            myList = Item
            first = False
        Else
            myList = myList & ", " & Item
        End If
    Next Item
    
    MakeCollectionString = myList
    
    Exit Function

ErrorHandler:
    ErrorMsg Err.Description, "MakeCollectionString"
End Function

Sub ErrorMsg(errMsg As String, funcName As String)
    Dim msg as String
    msg = "Error: " & errMsg & vbNewLine & "Skipped execution in " & funcName
    MsgBox msg, vbCritical, "Error"
    Debug.Print ("Error in " & funcName & ": " & errMsg)
    
End Sub
