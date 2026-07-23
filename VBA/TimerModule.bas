Attribute VB_Name = "TimerModule"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Digital countdown timer for VAY Bible Challenge (GitHub issue #2).
'
'    A generation counter invalidates any running countdown whenever the
'    timer is started, reset, or stopped, so stale loops exit cleanly.
'    DoEvents inside the wait loop keeps the slideshow responsive, so the
'    host can click the Correct/Incorrect score buttons while it runs.
'
'    Deck wiring (per .pptm, done in PowerPoint on Windows):
'      1. Import this module into the VBA project
'      2. Add a text box named "CountdownDisplay" to the Slide Master
'      3. Give a Start Timer shape the action: Run macro > StartCountdown
'
'    This module is distributed under the GNU General Public License v3.0
'    or later, like the rest of this project. See LICENSE.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const COUNTDOWN_SECONDS As Integer = 10
Private Const TIME_UP_LINGER_SECONDS As Integer = 3
Private Const DISPLAY_SHAPE As String = "CountdownDisplay"
Private Const TIME_UP_SOUND As String = "Time-Up.wav"

Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000

Private mGeneration As Long

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal sName As String, ByVal hMod As LongPtr, ByVal iFlags As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal sName As String, ByVal hMod As Long, ByVal iFlags As Long) As Long
#End If

' Assign to the Start Timer shape via Insert > Action > Run macro
Public Sub StartCountdown()
    mGeneration = mGeneration + 1
    Call RunCountdown(mGeneration)
End Sub

Public Sub ResetCountdown()
    ' Called on a wrong answer: the next buzzer gets a fresh countdown.
    ' Identical to StartCountdown today; kept separate so call sites
    ' document intent and the behaviors can diverge later (e.g. 5s rebuzz).
    Call StartCountdown
End Sub

Public Sub StopCountdown()
    mGeneration = mGeneration + 1
    Call SetDisplayText("")
End Sub

Private Sub RunCountdown(myGen As Long)
Dim secs As Integer
Dim tNext As Single
Dim startSlide As Long

On Error GoTo Bail
startSlide = ActivePresentation.SlideShowWindow.View.Slide.slideIndex

tNext = Timer
For secs = COUNTDOWN_SECONDS To 1 Step -1
    Call SetDisplayText(Trim(Str(secs)))
    tNext = tNext + 1
    Do While Timer < tNext
        DoEvents
        Sleep 20                                        ' keep CPU idle between polls
        If mGeneration <> myGen Then Exit Sub           ' reset/stop happened
        If Timer < tNext - 2 Then tNext = Timer + 1     ' midnight wrap of Timer
        If ActivePresentation.SlideShowWindow.View.Slide.slideIndex _
            <> startSlide Then GoTo Bail                ' host moved to another slide
    Loop
Next secs

Call SetDisplayText("TIME!")
Call PlayTimeUpSound

' Leave TIME! up briefly, then clear (unless a new countdown took over)
tNext = Timer + TIME_UP_LINGER_SECONDS
Do While Timer < tNext
    DoEvents
    Sleep 20
    If mGeneration <> myGen Then Exit Sub
Loop
Call SetDisplayText("")
Exit Sub

Bail:
If mGeneration = myGen Then Call SetDisplayText("")
End Sub

Private Sub PlayTimeUpSound()
Dim sPath As String

' The sound file lives next to the presentation, like the other music files
On Error Resume Next
sPath = ActivePresentation.Path & "\" & TIME_UP_SOUND
If Dir(sPath) <> "" Then Call PlaySound(sPath, 0, SND_FILENAME Or SND_ASYNC)
End Sub

Private Sub SetDisplayText(sText As String)
Dim oShape As Shape

For Each oShape In ActivePresentation.SlideMaster.Shapes
    If oShape.Name = DISPLAY_SHAPE Then
        oShape.TextFrame.TextRange.Text = sText

        ' Work-around for screen not updating problem (same as scoreboards):
        If Val(Application.Version) >= 12 Then Call RefreshMe(oShape)

        Exit For
    End If
Next oShape
End Sub
