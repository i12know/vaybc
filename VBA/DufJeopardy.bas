Attribute VB_Name = "DufJeopardy"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Jeopardy template created by Kevin Dufendach.
'    Please feel free to use and distribute this file and/or modify it
'    for your own needs.
'
'    Contact information:
'    Kevin Dufendach
'    krd.public+Jeopardy@gmail.com
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Jeopardy template for PowerPoint (Not in any way endorsed or affiliated with
'    the Jeopardy Game Show).
'    Software Copyright 2009-2013 Kevin Dufendach
'
'    These macro module is a free module: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This module is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private Const PLAYER_COUNT = 8
Private Const CATEGORY_COUNT = 6
Private Const ROW_COUNT = 5
Private Const BOARD_TOP = 144
Private Const BOARD_LEFT = 18
Private Const TILE_H_SPACING = 114
Private Const TILE_V_SPACING = 78
Private Const DOLLAR_CHARACTER = "$"

Sub correctPlayer1()
    Call ChangeResponse(1, 1)
End Sub
Sub correctPlayer2()
    Call ChangeResponse(2, 1)
End Sub
Sub correctPlayer3()
    Call ChangeResponse(3, 1)
End Sub
Sub correctPlayer4()
    Call ChangeResponse(4, 1)
End Sub
Sub correctPlayer5()
    Call ChangeResponse(5, 1)
End Sub
Sub correctPlayer6()
    Call ChangeResponse(6, 1)
End Sub
Sub correctPlayer7()
    Call ChangeResponse(7, 1)
End Sub
Sub correctPlayer8()
    Call ChangeResponse(8, 1)
End Sub
Sub incorrectPlayer1()
    Call ChangeResponse(1, -1)
End Sub
Sub incorrectPlayer2()
    Call ChangeResponse(2, -1)
End Sub
Sub incorrectPlayer3()
    Call ChangeResponse(3, -1)
End Sub
Sub incorrectPlayer4()
    Call ChangeResponse(4, -1)
End Sub
Sub incorrectPlayer5()
    Call ChangeResponse(5, -1)
End Sub
Sub incorrectPlayer6()
    Call ChangeResponse(6, -1)
End Sub
Sub incorrectPlayer7()
    Call ChangeResponse(7, -1)
End Sub
Sub incorrectPlayer8()
    Call ChangeResponse(8, -1)
End Sub

Function getSlideValue()
Dim oSlides As Slides
Dim oObject As Object
Dim sText As String
Dim sValue As String
Dim nDollarLoc As Integer

sText = DOLLAR_CHARACTER + "0"

For Each oObject In ActivePresentation.SlideShowWindow.View.Slide.Shapes
    If oObject.Name = "slideValue" Then
        sText = oObject.TextFrame.TextRange.Text
        Exit For
    End If
Next oObject

nDollarLoc = InStr(1, sText, DOLLAR_CHARACTER)
sValue = VBA.Strings.Right(sText, Len(sText) - nDollarLoc)

getSlideValue = Val(sValue)

End Function

Function getSlideRow() As Integer
Dim sSlideName As String

sSlideName = ActivePresentation.SlideShowWindow.View.Slide.Name
getSlideRow = Val(VBA.Strings.Mid(sSlideName, 9, 1))

End Function

Sub setTimerDelay()

Dim oMaster As Master
Dim oShape As Shape
Dim oAnimationSequence As Effect

Dim k As Integer

Dim sDelayTime As String
Dim dDelayTime As Double
Dim i As Integer
Dim j As Integer
Dim n As Integer
                
Dim iThisNumber As Integer
                
sDelayTime = InputBox("Enter delay time")
dDelayTime = Val(sDelayTime)
    
For k = 1 To 2
    If k = 1 Then
        Set oMaster = ActivePresentation.SlideMaster
    Else
        Set oMaster = ActivePresentation.TitleMaster
    End If

    For Each oShape In oMaster.Shapes
        If Len(oShape.Name) >= 11 Then
            If Left(oShape.Name, 10) = "timerLight" Then
                iThisNumber = Val(Mid(oShape.Name, 11, 1))
                For j = 1 To oMaster.TimeLine.InteractiveSequences.Count
                    For n = 1 To oMaster.TimeLine.InteractiveSequences.Item(j).Count
                        If oMaster.TimeLine.InteractiveSequences.Item(j).Item(n).Shape.Name = oShape.Name Then
                            With oMaster.TimeLine.InteractiveSequences.Item(j).Item(n)
                                If .Exit = msoTrue Then
                                    .Timing.TriggerDelayTime = dDelayTime / 5 * iThisNumber
                                    Exit For
                                End If
                            End With
                        End If
                    Next n
                Next j
            End If
        End If
    Next oShape
Next k

End Sub

Private Sub resetAllSlideNames()
Dim oSlide As Slide

For Each oSlide In ActivePresentation.Slides
    oSlide.Name = "Slide" + Trim(Str(oSlide.slideIndex))
Next

End Sub

Private Sub resetSlideNames()
Dim startingSlide As Integer
Dim b As Integer
Dim c As Integer
Dim r As Integer

For b = 1 To getNumberOfBoardSlides
    startingSlide = getBoardSlide(b).slideIndex + 1
    For c = 0 To (CATEGORY_COUNT - 1)
        For r = 0 To (ROW_COUNT - 1)
            ActivePresentation.Slides(startingSlide + c * 10 + r * 2).Name = "SlideB" + Trim(Str(b)) + numberToLetter(c + 1) + Trim(Str(r + 1))
        Next r
    Next c
Next b
End Sub

Sub testing()
    MsgBox Str(letterToNumber(numberToLetter(1)))
End Sub
Private Function numberToLetter(thisNumber As Integer) As String
    ' note: "1" is A, "0" is not
    numberToLetter = Trim(VBA.Strings.Chr(64 + thisNumber))
End Function
Private Function letterToNumber(thisLetter As String) As Integer
    letterToNumber = VBA.Strings.Asc(thisLetter) - 64
End Function


Private Sub changeSlideName()
    With PowerPoint.ActiveWindow.Selection.SlideRange
        If .Count > 1 Then Exit Sub
        MsgBox .Name + " is the current slide name"
        
        .Name = "FinalJeopardyResponseSlide"
        MsgBox "Successfully changed to " + .Name
        
    End With
        
End Sub

Function getSlideColumn() As Integer
Dim sSlideName As String

sSlideName = ActivePresentation.SlideShowWindow.View.Slide.Name
getSlideColumn = letterToNumber(VBA.Strings.Mid(sSlideName, 8, 1))

End Function

Function getSlideColumnFromValueBox()
Dim oSlides As Slides
Dim oObject As Object
Dim sText As String
Dim sValue As String
Dim nDollarLoc As Integer

sText = DOLLAR_CHARACTER + "0"

For Each oObject In ActivePresentation.SlideShowWindow.View.Slide.Shapes
    If oObject.Name = "slideValue" Then
        sText = oObject.TextFrame.TextRange.Text
        Exit For
    End If
Next oObject

getSlideColumn = Val(VBA.Strings.Left(sText, 1))

End Function

Private Sub ChangeResponse(player As Integer, direction As Integer)
Dim iSlideValue As Integer
Dim oMaster As Master
Dim oShape As Shape
Dim sScoreBoard As String
Dim iColumn As Integer
Dim k As Integer, nDollarLoc As Integer, nValue As Integer, newValue As Integer
Dim sText As String

iSlideValue = getSlideValue
nValue = GetScore(player)
newValue = nValue + direction * iSlideValue
Call ChangeValue(player, newValue)

' Change visibility on board
Call hideTileOnBoard(getSlideColumn, getSlideRow)

' Append score to VAY's Audit Trail on the last slide
With ActivePresentation.Slides(ActivePresentation.Slides.Count)
    .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player " & player & sNameOfPlayer(player) & " @" & numberToLetter(getSlideColumn) & getSlideRow & ": " & (direction * iSlideValue) & " from " & nValue & " = " & newValue & Chr(13))
End With

' Go to the next slide if that's what's supposed to happen
If direction = 1 Then
    Call DefaultAction
End If


End Sub

Private Function getBoardIndex(Optional slideIndex As Integer) As Integer
' Reports which Jeopardy number, starting with "1"

Dim iBoardSlideCounter As Integer
Dim sSlide As Slide

If slideIndex = 0 Then
    slideIndex = ActivePresentation.SlideShowWindow.View.Slide.slideIndex
End If

getBoardIndex = 1
iBoardSlideCounter = 0
For Each sSlide In ActivePresentation.Slides
    ' Board slides start with the name "BoardSlide"
    If Left(sSlide.Name, 10) = "BoardSlide" Then
        iBoardSlideCounter = iBoardSlideCounter + 1
    End If
    If sSlide.slideIndex >= slideIndex Then
        If iBoardSlideCounter = 0 Then
            getBoardIndex = 1
        Else
            getBoardIndex = iBoardSlideCounter
        End If
        Exit Function
    End If
Next sSlide

End Function

Private Function getBoardSlide(Optional iWhichJeopardy As Integer) As Slide
Dim sBoardName As String

If iWhichJeopardy = 0 Then
    iWhichJeopardy = getBoardIndex
End If

sBoardName = "BoardSlide" + Trim(Str(iWhichJeopardy))
Set getBoardSlide = ActivePresentation.Slides(sBoardName)

End Function

Private Function getDailyDoubleSlide(Optional iWhichDDouble As Integer) As Slide
Dim sSlideName As String

If iWhichDDouble = 0 Then
    iWhichDDouble = getBoardIndex
End If

sSlideName = "DailyDouble" + Trim(Str(iWhichDDouble))
Set getDailyDoubleSlide = ActivePresentation.Slides(sSlideName)

End Function

Private Function GetScore(player As Integer)
Dim iSlideValue As Integer
Dim oMaster As Master
Dim oShape As Shape
Dim sScoreBoard As String
Dim iColumn As Integer
Dim k As Integer, nDollarLoc As Integer, currentValue As Integer, newValue As Integer
Dim sText As String

sScoreBoard = "ScoreBoard" + LTrim(RTrim(Str(player)))

For k = 1 To 2
    If k = 1 Then
        Set oMaster = ActivePresentation.SlideMaster
    Else
        Set oMaster = ActivePresentation.TitleMaster
    End If

    For Each oShape In oMaster.Shapes
        If oShape.Name = sScoreBoard Then
            sText = oShape.TextFrame.TextRange.Text
            nDollarLoc = InStr(1, sText, DOLLAR_CHARACTER)
            currentValue = Val(VBA.Strings.Right(sText, Len(sText) - nDollarLoc))
            
            Exit For
        End If
    Next oShape
Next k

GetScore = currentValue

End Function

Private Sub ChangeValue(player As Integer, newValue As Integer)
Dim oMaster As Master
Dim oShape As Shape
Dim sScoreBoard As String
Dim k As Integer

sScoreBoard = "ScoreBoard" + LTrim(RTrim(Str(player)))

For k = 1 To 2
    If k = 1 Then
        Set oMaster = ActivePresentation.SlideMaster
    Else
        Set oMaster = ActivePresentation.TitleMaster
    End If

    For Each oShape In oMaster.Shapes
        If oShape.Name = sScoreBoard Then
            oShape.TextFrame.TextRange.Text = DOLLAR_CHARACTER + LTrim(RTrim(Str(newValue)))
            
            ' Work-around for screen not updating problem:
            If Val(Application.Version) >= 12 Then Call RefreshMe(oShape)
            
            Exit For
        End If
    Next oShape
Next k
End Sub

Sub ClearThisSlidesTileAndGoToNext()
    Call hideTileOnBoard(getSlideColumn, getSlideRow)
    Call DefaultAction

End Sub
Private Sub hideTileOnBoard(iColumn As Integer, iSlideRow As Integer)
Dim oShape As Shape
Dim sCheckString As String

sCheckString = "c" + Trim(Str(iColumn)) + "r" + Trim(Str(iSlideRow)) + "00Text"

For Each oShape In getBoardSlide().Shapes
    If oShape.Name = sCheckString Then
        oShape.Visible = msoFalse
        
        ' Work-around for screen not updating problem:
        If Val(Application.Version) >= 12 Then Call RefreshMe(oShape)
        
        Exit For
    End If
Next oShape

End Sub

Private Function getNumberOfBoardSlides() As Integer
    getNumberOfBoardSlides = getBoardIndex(ActivePresentation.Slides.Count)
End Function

Private Function getSlide(Board As Integer, column As Integer, row As Integer) As Slide
    ' Column is "1" for "A"
    Set getSlide = ActivePresentation.Slides("SlideB" + Trim(Str(Board)) + numberToLetter(column) + Trim(Str(row)))
End Function

Sub resetAll()
Dim oShape As Shape
Dim player As Integer
Dim k As Integer
Dim MsgResult As Long
Dim oMaster
Dim sScoreBoard As String
Dim c As Integer
Dim r As Integer, n As Integer
Dim DDr As Integer
Dim DDc As Integer
Dim DDi As Integer
Dim sValue As String
Dim testSlide As Slide
Dim tempText As String
Dim leftMargin As Integer
Dim horizontalSpacing As Integer
Dim topMargin As Integer
Dim verticalSpacing As Integer

MsgResult = MsgBox("This will reset the board and each player's score. " _
        & "Are you sure you want to do this?", vbYesNo)

Randomize 'initialize random number generator

If MsgResult = vbYes Then
    ' n becomes a reference to the index of the board slide (i.e. n = 1 for single jeopardy, n = 2 for double, etc.)
    For n = 1 To getNumberOfBoardSlides
        On Error GoTo ErrorHandlerNextBoard
        
        With getBoardSlide(n)
            On Error GoTo 0
    
            For Each oShape In .Shapes
                oShape.Visible = msoTrue
            Next oShape
        
            For c = 1 To CATEGORY_COUNT
                For r = 1 To ROW_COUNT
                    On Error GoTo ErrorHandlerContinueToNextShape
                    sValue = .Shapes("c" + Trim(Str(c)) + "r" + Trim(Str(r)) + "00Text").TextFrame.TextRange.Text
                    getSlide(n, c, r).Shapes("slideValue").TextFrame.TextRange.Text = sValue

ContinueToNextShape:
                Next
            Next
            
            ' use first shape to get margin for calculations
            Set oShape = .Shapes("c1r100Text")
            leftMargin = oShape.Left
            topMargin = oShape.Top
            
            ' use shape to the right and below first shape get spacing between shapes
            horizontalSpacing = .Shapes("c2r100Text").Left - leftMargin
            verticalSpacing = .Shapes("c1r200Text").Top - topMargin
            
            'Daily Double
            ' DailyDouble slides are named "DailyDouble1" for Reg Jeopardy,
            ' "DailyDouble2" and "DailyDouble3" for double Jeopardy, etc.
            
            ' Create a dictionary to store Daily Double locations
            Dim DailyDoubleLocations As Object
            Set DailyDoubleLocations = CreateObject("Scripting.Dictionary")
            
            For k = 1 To n
                DDi = DDi + 1
                'Set testSlide = getDailyDoubleSlide(DDi)
                
                If ActivePresentation.Tags("EnableDailyDoubles") = "true" Then
                    Do
                        DDc = Int(Rnd * (CATEGORY_COUNT - 1)) ' VAY used only 5 columns
                        DDr = Int(Rnd * ROW_COUNT)
                    Loop While DailyDoubleLocations.Exists(DDc & "_" & DDr)  ' Ensure unique location
                    ' Implemented by do-loop-while above: check to be sure the same location is not used if there are multiple daily doubles
                    
                    ' Add this location to the dictionary
                    DailyDoubleLocations.Add (DDc & "_" & DDr), 1

                    With getDailyDoubleSlide(DDi)
                            .Tags.Add "DailyDoubleColumn", Str(DDc)
                            .Tags.Add "DailyDoubleRow", Str(DDr)
                            .Tags.Add "DailyDoubleSlideNumber", Str(DDi)
                    End With
                    .Shapes("DailyDoubleLocationBox" + Trim(Str(k))).TextFrame.TextRange.Text = numberToLetter(DDc + 1) + Trim(Str(DDr + 1))
                    
                    With .Shapes("DailyDoubleSender" + Trim(Str(k)))
                        .Top = topMargin + DDr * verticalSpacing
                        .Left = leftMargin + DDc * horizontalSpacing
                    End With
                
                Else
                    With .Shapes("DailyDoubleSender" + Trim(Str(k)))
                        .Top = -verticalSpacing - 10
                        .Left = leftMargin
                    End With
                    .Shapes("DailyDoubleLocationBox" + Trim(Str(k))).TextFrame.TextRange.Text = "off"
                End If
            Next k
        End With 'getBoardSlide(n)
        
ContinueToNextBoard:
        On Error GoTo 0
    Next n
    
    For player = 1 To PLAYER_COUNT
        Call ChangeValue(player, 0)
    Next player
    
    
    ActivePresentation.Tags.Add "Wager1", "0"
    ActivePresentation.Tags.Add "Wager2", "0"
    ActivePresentation.Tags.Add "Wager3", "0"
    ActivePresentation.Tags.Add "Wager4", "0"
    ActivePresentation.Tags.Add "Wager5", "0"
    ActivePresentation.Tags.Add "Wager6", "0"
    ActivePresentation.Tags.Add "Wager7", "0"
    ActivePresentation.Tags.Add "Wager8", "0"
    
    Slide68.WagerBox1.Text = ""
    Slide68.WagerBox2.Text = ""
    Slide68.WagerBox3.Text = ""
    Slide68.WagerBox4.Text = ""
    Slide68.WagerBox5.Text = ""
    Slide68.WagerBox6.Text = ""
    Slide68.WagerBox7.Text = ""
    Slide68.WagerBox8.Text = ""
End If

' Reset VAY's Audit Trail on the last slide
With ActivePresentation.Slides(ActivePresentation.Slides.Count)
    .Shapes.Placeholders(2).TextFrame.TextRange.Text = ""
End With

Exit Sub

ErrorHandlerContinueToNextShape:
Resume ContinueToNextShape

Exit Sub
ErrorHandlerNextBoard:
Resume ContinueToNextBoard

End Sub

Public Sub SetDailyDoubleScore()
Dim x
Dim sWager As String
Dim DDc As Integer
Dim DDr As Integer
Dim iWager As Integer
Dim oShape As Shape
Dim boardSlide As Slide
Dim sDailyDoubleSlide As Slide
Dim boardIndex As Integer
Dim dailyDoubleIndex As Integer
Dim slideObject As Object

'' MsgBox ("Entering Daily Double Public Sub.")

boardIndex = getBoardIndex
Set boardSlide = getBoardSlide(boardIndex)

'Set sDailyDoubleSlide = ActivePresentation.Slides("DailyDouble" + Trim(Str(boardIndex)))
Set sDailyDoubleSlide = ActivePresentation.SlideShowWindow.View.Slide

'Set x = ActivePresentation.Slides("DailyDouble").Shapes("WagerBox")
With ActivePresentation.SlideShowWindow.View.Slide
    'sWager = .Shapes("WagerBox").Value
    dailyDoubleIndex = Int(Val(.Tags("DailyDoubleSlideNumber")))
    
    If dailyDoubleIndex = 1 Then
        sWager = Slide67.WagerBox.Value
    ElseIf dailyDoubleIndex = 2 Then
        sWager = Slide132.WagerBox.Value
    ElseIf dailyDoubleIndex = 3 Then
        sWager = Slide133.WagerBox.Value
    Else
        MsgBox ("Daily Double did not work correctly.")
        Exit Sub
    End If
    'sWager = sDailyDoubleSlide.Shapes("WagerBox").Value

    iWager = Int(Val(sWager))

    DDc = Int(Val(.Tags("DailyDoubleColumn")))
    DDr = Int(Val(.Tags("DailyDoubleRow")))
End With

With getSlide(boardIndex, DDc + 1, DDr + 1)
    For Each oShape In .Shapes
        If oShape.Name = "slideValue" Then
            oShape.TextFrame.TextRange.Text = DOLLAR_CHARACTER + Trim(Str(iWager))
            Exit For
        End If
    Next
End With


ActivePresentation.SlideShowWindow.View.GotoSlide getSlide(boardIndex, DDc + 1, DDr + 1).slideIndex

End Sub

Private Sub CreateBoardAndSlides()
Dim oSlide As Slide
Dim oBoardSlide As Slide
Dim oShape As Shape
Dim tempText As String
Dim kRow As Integer
Dim kColumn As Integer

' Create board slide
Set oBoardSlide = ActivePresentation.Slides.Add(4, ppLayoutBlank)
oBoardSlide.Name = "BoardSlide1"

' Create question and answer slides
For kColumn = 1 To CATEGORY_COUNT
    With oBoardSlide.Shapes.AddShape(msoShapeRoundedRectangle, 18 + (kColumn - 1) * 114, 54, 108, 54)
        .TextFrame.TextRange.Text = "Category " & Trim(Str(kColumn))
        .Fill.BackColor.SchemeColor = ppBackground
        .Fill.ForeColor.SchemeColor = ppBackground
        .Line.Visible = msoFalse
        .Name = "Category" & Trim(Str(kColumn))
    End With
    
    
    For kRow = 1 To ROW_COUNT
        ' Create question slide
        Set oSlide = ActivePresentation.Slides.Add(3 + ((kColumn - 1) * ROW_COUNT + kRow) * 2, ppLayoutTitle)
        
        oSlide.Shapes.Placeholders.Item(1).TextFrame.TextRange = "Category " & LTrim(RTrim(Str(kColumn))) & ", " & DOLLAR_CHARACTER & LTrim(RTrim(Str(kRow))) & "00 question"
        
        Set oShape = oSlide.Shapes.AddShape(msoShapeRoundedRectangle, 642, 0, 78, 36)
        oShape.TextFrame.TextRange.Text = LTrim(RTrim(Str(kColumn))) & ": " & DOLLAR_CHARACTER & LTrim(RTrim(Str(kRow))) & "00"
        oShape.Name = "slideValue"
        
        ' Create background on Board
        With oBoardSlide.Shapes.AddShape(msoShapeRoundedRectangle, 18 + (kColumn - 1) * 114, 132 + (kRow - 1) * 78, 108, 72)
            
            '.Select
'            .TextFrame.TextRange.Text = "$" & Trim(Str(kRow)) & "00"
            .Fill.BackColor.SchemeColor = ppBackground
            .Fill.ForeColor.SchemeColor = ppBackground
            .Line.Visible = msoFalse
            'TempText = Trim(Str(kColumn)) & "-" & Trim(Str(kRow)) & "00Back"
            tempText = "c" & Trim(Str(kColumn)) & "r" & Trim(Str(kRow)) & "00Back"
            .Name = tempText
            
            
            '.ActionSettings
        End With
        ' Create dollar amount text on Board
        With oBoardSlide.Shapes.AddShape(msoShapeRectangle, 18 + (kColumn - 1) * 114, 138 + (kRow - 1) * 78, 108, 60)
            .TextFrame.TextRange.Text = DOLLAR_CHARACTER & Trim(Str(kRow)) & "00"
            .Fill.BackColor.SchemeColor = ppBackground
            .Fill.ForeColor.SchemeColor = ppBackground
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            'TempText = Trim(Str(kcolumn)) & "-" & Trim(Str(krow)) & "00Text"
            tempText = "c" & Trim(Str(kColumn)) & "r" & Trim(Str(kRow)) & "00Text"
            .Name = tempText
        
            .ActionSettings.Item(ppMouseClick).Action = ppActionHyperlink
            tempText = oSlide.SlideID & "," & oSlide.slideIndex & "," & oSlide.Shapes.Title.TextFrame.TextRange.Text
            .ActionSettings.Item(ppMouseClick).Hyperlink.SubAddress = tempText
        End With
        
        ' Create answer slide
        With ActivePresentation.Slides.Add(3 + ((kColumn - 1) * ROW_COUNT + kRow) * 2, ppLayoutText)
            .Shapes.Placeholders.Item(1).TextFrame.TextRange = "Category " & LTrim(RTrim(Str(kColumn))) & ", " & DOLLAR_CHARACTER & LTrim(RTrim(Str(kRow))) & "00 answer"
        
            With .Shapes.AddShape(msoShapeRoundedRectangle, 642, 0, 78, 36)
                .TextFrame.TextRange.Text = LTrim(RTrim(Str(kColumn))) & ": " & DOLLAR_CHARACTER & LTrim(RTrim(Str(kRow))) & "00"
                .Name = "slideValue"
            End With
        End With
    Next
Next


End Sub

Private Sub DefaultAction()
Dim tempText As String
Dim boardSlideIndex As Integer

    tempText = ActivePresentation.Tags("ReturnToBoard")
    
    If tempText = "true" Then
        SlideShowWindows(1).View.GotoSlide getBoardSlide.slideIndex
    Else
        SlideShowWindows(1).View.Next
    End If
End Sub

Public Sub ImportFromExcel()
Dim ExcelApp As Excel.Application
Dim thisAnswer As String, thisQuestion As String, thisComments As String, thisCategoryText As String
Dim CatCounter As Integer, ValCounter As Integer, lngCount As Integer
Dim thisSheet As Worksheet
Dim thisWorkbookName As String
Dim thisWorkbook As Excel.Workbook
Dim boardSlideIndex As Integer
Dim thisObject As Object
Dim thisValue As Integer
Dim n As Integer
Dim iNextSlideIndex As Integer
Dim rowStart As Integer

On Error GoTo ErrorHandler

With Application.FileDialog(msoFileDialogOpen)
    .AllowMultiSelect = False
    .Filters.Add "All Microsoft Office Excel Files", "*.xl*; *.xls; *.xla; *.xlt; *.xlm; *.xlc; *.xlw; *.htm; *.html; *.mht; *.mhtml; *.odc; *.uxdc; *.xlsx; *.xlsm; *.xltx; *.xltm", 1
    .FilterIndex = 1
    
    .Show
    
    ' Display paths of each file selected
    For lngCount = 1 To .SelectedItems.Count
        thisWorkbookName = .SelectedItems(lngCount)
    Next lngCount
End With

If thisWorkbookName = "" Then Exit Sub

Set ExcelApp = Excel.Application
Set thisWorkbook = ExcelApp.Workbooks.Open(thisWorkbookName)
On Error GoTo ErrorHandlerCloseFile

Set thisSheet = thisWorkbook.Worksheets("JeopardyQuestionTemplate")

For n = 1 To getNumberOfBoardSlides
    rowStart = 4 + (n - 1) * 31
    
    'Find out where to start (where the "board" is)
    boardSlideIndex = ActivePresentation.Slides("BoardSlide" + Trim(Str(n))).slideIndex
        
    For CatCounter = 0 To CATEGORY_COUNT - 1
        thisCategoryText = thisSheet.Range("A" & Trim(Str(CatCounter * ROW_COUNT + rowStart)))
        With ActivePresentation.Slides(boardSlideIndex)
            'Category1 = sample name of one of the category boxes
            Set thisObject = .Shapes("Category" & Trim(Str(CatCounter + 1)))
            thisObject.TextFrame.TextRange.Text = thisCategoryText
        End With
        For ValCounter = 0 To ROW_COUNT - 1
            ' Get Results from Spreadsheet
            thisValue = Int(Val(thisSheet.Range("B" & Trim(Str(ValCounter + CatCounter * ROW_COUNT + rowStart)))))
            
            thisAnswer = thisSheet.Range("C" & Trim(Str(ValCounter + CatCounter * ROW_COUNT + rowStart)))
            thisQuestion = thisSheet.Range("D" & Trim(Str(ValCounter + CatCounter * ROW_COUNT + rowStart)))
            thisComments = thisSheet.Range("E" & Trim(Str(ValCounter + CatCounter * ROW_COUNT + rowStart)))
            
            'Put Results into proper boxes
            getBoardSlide(n).Shapes("c" + Trim(Str(CatCounter + 1)) + "r" + Trim(Str(ValCounter + 1)) + "00Text").TextFrame.TextRange.Text = DOLLAR_CHARACTER + Trim(Str(thisValue))
            With getSlide(n, CatCounter + 1, ValCounter + 1)
                .Shapes.Placeholders(1).TextFrame.TextRange.Text = Replace(thisAnswer, "~", Chr(13) + "=======" + Chr(13)) 'split string w/ new line for thisAnswer
                .Shapes("slideValue").TextFrame.TextRange.Text = DOLLAR_CHARACTER + Trim(Str(thisValue))
                iNextSlideIndex = .slideIndex + 1
            End With
            
            With ActivePresentation.Slides(iNextSlideIndex)
                .Shapes.Placeholders(1).TextFrame.TextRange.Text = Replace(thisQuestion, "~", Chr(13) + "=======" + Chr(13)) 'split string w/ new line for thisQuestion
                .Shapes.Placeholders(2).TextFrame.TextRange.Text = Replace(thisComments, "~", Chr(13) + "=======" + Chr(13)) 'split string w/ new line for thisComments
                .Shapes("slideValue").TextFrame.TextRange.Text = DOLLAR_CHARACTER + Trim(Str(thisValue))
            End With
            
        Next ValCounter
    Next CatCounter
Next n

On Error GoTo ErrorHandler
thisWorkbook.Close False
Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
ErrorHandlerCloseFile:
thisWorkbook.Close False

' Error handler (goes here before workbook opened)
ErrorHandler:
MsgBox ("There was an error in importing this workbook")

End Sub



Public Sub goToBoardSlide()
    ActivePresentation.SlideShowWindow.View.GotoSlide getBoardSlide.slideIndex
End Sub

Private Sub GetObjectName()

With ActiveWindow.Selection
    MsgBox .ShapeRange(1).Name
End With


End Sub

Public Sub AdjustScores()
Dim i As Integer

With AdjustScoresBox
    .TextBox1.Value = GetScore(1)
    .TextBox2.Value = GetScore(2)
    .TextBox3.Value = GetScore(3)
    .TextBox4.Value = GetScore(4)
    .TextBox5.Value = GetScore(5)
    .TextBox6.Value = GetScore(6)
    .TextBox7.Value = GetScore(7)
    .TextBox8.Value = GetScore(8)
    
    
    .Show
    
    If .Tag = "OK" Then
    
        ' Append score to VAY's Audit Trail on the last slide
        With ActivePresentation.Slides(ActivePresentation.Slides.Count)
            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 1" & sNameOfPlayer(1) & " @AdjustScores: " & AdjustScoresBox.TextBox1.Value & " from " & GetScore(1) & Chr(13))
            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 2" & sNameOfPlayer(2) & " @AdjustScores: " & AdjustScoresBox.TextBox2.Value & " from " & GetScore(2) & Chr(13))
            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 3" & sNameOfPlayer(3) & " @AdjustScores: " & AdjustScoresBox.TextBox3.Value & " from " & GetScore(3) & Chr(13))
'            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 4 @AdjustScores: " & AdjustScoresBox.TextBox4.Value & " from " & GetScore(4) & Chr(13))
'            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 5 @AdjustScores: " & AdjustScoresBox.TextBox5.Value & " from " & GetScore(5) & Chr(13))
'            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 6 @AdjustScores: " & AdjustScoresBox.TextBox6.Value & " from " & GetScore(6) & Chr(13))
'            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 7 @AdjustScores: " & AdjustScoresBox.TextBox7.Value & " from " & GetScore(7) & Chr(13))
'            .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player 8 @AdjustScores: " & AdjustScoresBox.TextBox8.Value & " from " & GetScore(8) & Chr(13))
        End With
    
        Call ChangeValue(1, Val(.TextBox1.Value))
        Call ChangeValue(2, Val(.TextBox2.Value))
        Call ChangeValue(3, Val(.TextBox3.Value))
        Call ChangeValue(4, Val(.TextBox4.Value))
        Call ChangeValue(5, Val(.TextBox5.Value))
        Call ChangeValue(6, Val(.TextBox6.Value))
        Call ChangeValue(7, Val(.TextBox7.Value))
        Call ChangeValue(8, Val(.TextBox8.Value))
    End If
End With

End Sub

Public Sub GoToFinalJeopardyResponseSlide()
Dim oTextBox As TextBox

    ActivePresentation.Tags.Add "Wager1", Slide68.WagerBox1.Text
    ActivePresentation.Tags.Add "Wager2", Slide68.WagerBox2.Text
    ActivePresentation.Tags.Add "Wager3", Slide68.WagerBox3.Text
    ActivePresentation.Tags.Add "Wager4", Slide68.WagerBox4.Text
    ActivePresentation.Tags.Add "Wager5", Slide68.WagerBox5.Text
    ActivePresentation.Tags.Add "Wager6", Slide68.WagerBox6.Text
    ActivePresentation.Tags.Add "Wager7", Slide68.WagerBox7.Text
    ActivePresentation.Tags.Add "Wager8", Slide68.WagerBox8.Text

ActivePresentation.SlideShowWindow.View.GotoSlide ActivePresentation.Slides("FinalJeopardyResponseSlide").slideIndex

End Sub

Private Sub ChangeFinalResponse(iPlayer As Integer, iDirection As Integer)
Dim sWager As String
Dim iWager As Integer
Dim iValue As Integer
Dim newValue As Integer

sWager = ActivePresentation.Tags.Item("Wager" + Trim(Str(iPlayer)))
iWager = Int(Val(sWager))

iValue = GetScore(iPlayer)
newValue = iValue + iDirection * iWager
Call ChangeValue(iPlayer, newValue)

' Append score to VAY's Audit Trail on the last slide
With ActivePresentation.Slides(ActivePresentation.Slides.Count)
    .Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter ("Player " & iPlayer & sNameOfPlayer(iPlayer) & " @Final: " & (iDirection * iWager) & " From " & iValue & " = " & newValue & Chr(13))
End With

End Sub

Sub correctFinalPlayer1()
    Call ChangeFinalResponse(1, 1)
End Sub
Sub correctFinalPlayer2()
    Call ChangeFinalResponse(2, 1)
End Sub
Sub correctFinalPlayer3()
    Call ChangeFinalResponse(3, 1)
End Sub
Sub correctFinalPlayer4()
    Call ChangeFinalResponse(4, 1)
End Sub
Sub correctFinalPlayer5()
    Call ChangeFinalResponse(5, 1)
End Sub
Sub correctFinalPlayer6()
    Call ChangeFinalResponse(6, 1)
End Sub
Sub correctFinalPlayer7()
    Call ChangeFinalResponse(7, 1)
End Sub
Sub correctFinalPlayer8()
    Call ChangeFinalResponse(8, 1)
End Sub
Sub incorrectFinalPlayer1()
    Call ChangeFinalResponse(1, -1)
End Sub
Sub incorrectFinalPlayer2()
    Call ChangeFinalResponse(2, -1)
End Sub
Sub incorrectFinalPlayer3()
    Call ChangeFinalResponse(3, -1)
End Sub
Sub incorrectFinalPlayer4()
    Call ChangeFinalResponse(4, -1)
End Sub
Sub incorrectFinalPlayer5()
    Call ChangeFinalResponse(5, -1)
End Sub
Sub incorrectFinalPlayer6()
    Call ChangeFinalResponse(6, -1)
End Sub
Sub incorrectFinalPlayer7()
    Call ChangeFinalResponse(7, -1)
End Sub
Sub incorrectFinalPlayer8()
    Call ChangeFinalResponse(8, -1)
End Sub

Sub RefreshMe(oSh As Shape)
    
    Dim Offset As Single
    Offset = ActivePresentation.PageSetup.SlideHeight + 10
    oSh.Top = oSh.Top + Offset
    DoEvents
    oSh.Top = oSh.Top - Offset

End Sub

Sub RenameShape()
    Dim sName As String
    Dim i As Integer
                
    With ActiveWindow.Selection.ShapeRange
        If .Count = 0 Then
            'message ("No appropriate object selected")
            Exit Sub
        ElseIf .Count = 1 Then
            sName = InputBox("Enter a shape name (currently " & .Item(1).Name & ")")
            If sName = "" Then Exit Sub
            .Item(1).Name = sName
        Else
            sName = InputBox("Enter a shape name")
            If sName = "" Then Exit Sub
            
            For i = 1 To .Count
                .Item(i).Name = sName + Trim(Str(i))
            Next i
        End If
    End With
End Sub

Sub RenameSlide()
    Dim sName As String
    Dim i As Integer
                
    With ActiveWindow.Selection.SlideRange
        If .Count = 0 Then
            'message ("No appropriate object selected")
            Exit Sub
        ElseIf .Count = 1 Then
            sName = InputBox("Enter a slide name (currently " & .Item(1).Name & ")")
            If sName = "" Then Exit Sub
            .Item(1).Name = sName
        Else
            sName = InputBox("Enter a slide name (name1, name2, ..., nameN)")
            If sName = "" Then Exit Sub
            
            For i = 1 To .Count
                .Item(i).Name = sName + Trim(Str(i))
            Next i
        End If
    End With
End Sub


' Attempt by Bumble from VAY July 2023
'' This code is used to retrieve the Team Names from the Master Slide Templates to use them later on
'' by overloading the initial slide showing event
Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
    
    If SSW.View.CurrentShowPosition = _
        SSW.Presentation.SlideShowSettings.startingSlide Then
                
        Dim oLayout As CustomLayout
        Dim oShape As Shape
    
        ' Assume that the Slide Master you want to access is the first one
        Set oLayout = ActivePresentation.SlideMaster.CustomLayouts(1)
        
        ' Loop through all shapes on the Slide Master
        For Each oShape In oLayout.Shapes
            Debug.Print oShape.Name
            ' Check if the shape is a TextBox
            If oShape.Type = msoTextBox Then
                ' Perform actions with the TextBox.
                ' For example, change the text:
                ' ' oShape.TextFrame.TextRange.Text = "New text"
'' worked                MsgBox oShape.TextFrame.TextRange.Text
            End If
        Next oShape
'' <failed        MsgBox oLayout.Shapes("Team1").TextFrame.TextRange.Text
''        MsgBox oLayout.Shapes("TextBox 2").TextFrame.TextRange.Text
               
        ' Initialize your string variable
        ' ' Dim team1Name As String

        ' Assume that the Slide Master you want to access is the first one
        ' ' Set masterLayout = ActivePresentation.SlideMaster.CustomLayouts(1)

        ' Access the TextBox on the Slide Master
        ' ' team1Name = masterLayout.Shapes("Team1").TextFrame.TextRange.Text

        ' Display the value in a message box
        ' ' MsgBox "The value of team1Name is: " & team1Name

        ' ' MsgBox "First slide in the slide show with " & team1Name
'' failed/>
    End If
End Sub


Private Function sNameOfPlayer(playerNo As Integer) As String
        Dim oLayout As CustomLayout
        Set oLayout = ActivePresentation.SlideMaster.CustomLayouts(1)
    
    Select Case playerNo
        Case 1
            sNameOfPlayer = oLayout.Shapes("TextBox 2").TextFrame.TextRange.Text
        Case 2
            sNameOfPlayer = oLayout.Shapes("TextBox 3").TextFrame.TextRange.Text
        Case 3
            sNameOfPlayer = oLayout.Shapes("TextBox 4").TextFrame.TextRange.Text
        Case Else
            sNameOfPlayer = "UNK"
    End Select
End Function

