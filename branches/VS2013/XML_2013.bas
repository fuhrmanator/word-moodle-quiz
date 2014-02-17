Attribute VB_Name = "XML"
'MS WORD TEMPLATE FOR MAKING MOODLE QUIZES
'=================================================
'(uses XML interchange format instead of GIFT)
'by Vyatcheslav Yatskovsky (yatskovsky@gmail.com)

'CREDITS
'based on the GIFTconverter template by Mikko Rusama
'inspired by OpenOffice aesthetical version by Enrique Castro

'CHANGE LOG
'4 Apr 2006   moved to UTF8
'23 Apr 2006  image support added
'24 Nov 2006  small improvements
'April 2012 - amalgamated v20 and added in changes from Luckas v12 (Essay), New XML based English toolbar - Lael Grant

'TO DO:
'!How convert (succesfully) to dotm template? currently ends with macros not recognised and freeze on creation of document from template
'!Why isn't textrange.xml read when creating from Word 2010 template?

'+ Images are imported, but are not named correctly, and therefore don't link
'   first image straight name, second = imagename_1.png _2.png etc.
'   Image link shown: http://moodle.wcc.nsw.edu.au/file.php/478/E%3Amoodledata/478/Images_forQuizQuestions/WriteQuestionCategoryHere
'+ Missing Word, Essay, True & False questions all end up with the header content
'  appended to the end of their question text if they are the last question in the quiz...
'+ Be nice to be able to specify the Question Grade & Penalty Grades
'+ All feedback is generated as <generalfeedback> rather than <feedback> which ends mcq's etc. Need to fix.
'+ update to support 2.2 and provide separate macro for exporting to 1.9 / 2.2

'************************
'The MIT License
'Copyright (c) 2005 Mikko Rusama, 2006 Vyatcheslav Yatskovky
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'************************

' General purpose styles.
Const STYLE_NORMAL = wdStyleNormal

Const STYLE_FEEDBACK = "Feedback"
Const STYLE_ANSWERWEIGHT = "AnswerWeight"

Const STYLE_SHORTANSWERQ = "Q Short Answer"
Const STYLE_MULTICHOICEQ = "Q Multi Choice"
Const STYLE_MATCHINGQ = "Q Matching"
Const STYLE_NUMERICALQ = "Q Numerical"
Const STYLE_MISSINGWORDQ = "Q Missing Word"
Const STYLE_TRUESTATEMENT = "TrueStatement"
Const STYLE_FALSESTATEMENT = "FalseStatement"
Const STYLE_CORRECTANSWER = "Correct Answer"
Const STYLE_INCORRECTANSWER = "Incorrect Answer"
Const STYLE_SHORT_ANSWER = "Short Answer"
Const STYLE_LEFT_PAIR = "LeftPair"
Const STYLE_RIGHT_PAIR = "RightPair"
Const STYLE_BLANK_WORD = "BlankWord"
Const STYLE_COMMENT = "Comment"
' Supplement(ed by) Daniel
Const STYLE_MULTICHOICEQ_FIXANSWER = "Q Multi Choice FixAnswer"
Const STYLE_MATCHINGQ_FIXANSWER = "Q Matching FixAnswer"
Const STYLE_NUM_TOLERANCE = "Num Tolerance"
Const STYLE_QUESTIONNAME = "Questionname"
'from v12
Const STYLE_ESSAY = "Q Essay"

' saves the current question type
Dim questionType As String
Dim xmlpath As String

' Prefix for the filename
Const FILE_PREFIX = "Moodle_Questions_" 'this has an underscore obscured by the line

' Add an Essay
Public Sub AddEssay(control As IRibbonControl)
    AddParagraphOfStyle STYLE_ESSAY, "Insert An Essay question here (an Open Question). [This can not be the last question in the document.]"
End Sub

' Add Multiple Choice Question to the end of the active document
Public Sub AddMultipleChoiceQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_MULTICHOICEQ, "Insert Multiple Choice Question"
    
End Sub

' Add Matching Question to the end of the active document
Sub AddMatchingQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_MATCHINGQ, "Insert Matching Question"
End Sub

' Add Numerical Question to the end of the active document
Sub AddNumericalQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_NUMERICALQ, "Insert Numerical Question"
End Sub


' Add Short Answer Question to the end of the active document
Sub AddShortAnswerQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_SHORTANSWERQ, "Insert Short Answer Question"
End Sub

' Add Missing Word Question
Sub AddMissingWordQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_MISSINGWORDQ, "Insert Missing Word Question. Then select the missing word!"
End Sub

' Marks the blank word
Public Sub MarkBlankWord(control As IRibbonControl)

    Set aRange = ActiveDocument.Range(Start:=Selection.Words(1).Start, End:=Selection.Words(1).End)
    If Selection.Words(1).Style = STYLE_BLANK_WORD Then
        aRange.Select
        Selection.Find.ClearFormatting
    Else
        'RTrim(ActiveDocument.Words(1)).Style = STYLE_BLANK_WORD
        aRange.Style = STYLE_BLANK_WORD
    End If
End Sub

' Add feedback - this doesn't seem to work in options or between options. Only creates <generalfeedback> rather than <feedback> and cuts off the rest of the text.
Sub AddQuestionFeedback(control As IRibbonControl)
    If Selection.Range.Style = STYLE_ANSWERWEIGHT Or _
       Selection.Range.Style = STYLE_SHORTANSWERQ Or _
       Selection.Range.Style = STYLE_MISSINGWORDQ Or _
       Selection.Range.Style = STYLE_CORRECTANSWER Or _
       Selection.Range.Style = STYLE_INCORRECTANSWER Or _
       Selection.Range.Style = STYLE_SHORT_ANSWER Or _
       Selection.Range.Style = STYLE_RIGHT_PAIR Or _
       Selection.Range.Style = STYLE_NUM_TOLERANCE Or _
       Selection.Range.Style = STYLE_TRUESTATEMENT Or _
       Selection.Range.Style = STYLE_FALSESTATEMENT Or _
       Selection.Range.Style = STYLE_QUESTIONNAME Or _
       Selection.Range.Style = STYLE_BLANK_WORD Then
        InsertAfterRange "Insert feedback of the previous choice or answer here.", _
                         STYLE_FEEDBACK, Selection.Paragraphs(1).Range
                               
                        
    Else 'Error: Give Instructions:
        MsgBox "Feedback is placed at the end of the last possible response. " & vbCr & _
               "It doesn't work for True/False questions." & vbCr & _
               "Place the cursor on top of the question or answer you are giving feedback for.", vbExclamation
        
    End If
End Sub
' Add tolerance
Sub AddNumericalTolerance(control As IRibbonControl)
    If Selection.Range.Style = STYLE_SHORT_ANSWER Then
        InsertAfterRange "Replace me with Tolerance for the answer as a Decimal. Eg: 0.01", _
                         STYLE_NUM_TOLERANCE, Selection.Paragraphs(1).Range
    Else 'Error: Give Instructions:
        MsgBox " " & vbCr & _
               "Place the cursor at the end of the numerical answer.", vbExclamation
    End If
End Sub

' Add QuestionName / Question Title
Sub AddQuestionTitle(control As IRibbonControl)
    If Selection.Range.Style = STYLE_ANSWERWEIGHT Or _
       Selection.Range.Style = STYLE_SHORTANSWERQ Or _
       Selection.Range.Style = STYLE_MISSINGWORDQ Or _
       Selection.Range.Style = STYLE_CORRECTANSWER Or _
       Selection.Range.Style = STYLE_NUM_TOLERANCE Or _
       Selection.Range.Style = STYLE_INCORRECTANSWER Or _
       Selection.Range.Style = STYLE_TRUESTATEMENT Or _
       Selection.Range.Style = STYLE_SHORT_ANSWER Or _
       Selection.Range.Style = STYLE_FALSESTATEMENT Or _
       Selection.Range.Style = STYLE_RIGHT_PAIR Or _
       Selection.Range.Style = STYLE_BLANK_WORD Then
        InsertAfterRange "Add a question title.", _
             STYLE_QUESTIONNAME, Selection.Paragraphs(1).Range
    Else 'Error: Give Instructions:
        MsgBox "Feedback to insert at the end of the last response selected. " & vbCr & _
               "The title must appear before the feedback" & vbCr & _
               "Place the cursor at the end of the last line selected", vbExclamation
    End If
End Sub


' Add a true statement of the true-false question
Sub AddTrueStatement(control As IRibbonControl)
    AddParagraphOfStyle STYLE_TRUESTATEMENT, "True-false question: insert a TRUE statement here (not at the end of the document)"
End Sub

' Add a false statement of the true-false question
Sub AddFalseStatement(control As IRibbonControl)
    AddParagraphOfStyle STYLE_FALSESTATEMENT, "True-false question: insert a FALSE statement here (not at the end of the document)"
End Sub

' Add a comment
Sub AddComment(control As IRibbonControl)
    AddParagraphOfStyle STYLE_COMMENT, ""
End Sub
Sub PasteImage(control As IRibbonControl)
'  Adds an image from the clipboard into a question.

    If Selection.Range.Style = STYLE_SHORTANSWERQ Or _
       Selection.Range.Style = STYLE_MISSINGWORDQ Or _
       Selection.Range.Style = STYLE_MULTICHOICEQ Or _
       Selection.Range.Style = STYLE_MATCHINGQ Or _
       Selection.Range.Style = STYLE_NUMERICALQ Or _
       Selection.Range.Style = STYLE_TRUESTATEMENT Or _
       Selection.Range.Style = STYLE_FALSESTATEMENT Or _
       Selection.Range.Style = STYLE_MULTICHOICEQ_FIXANSWER Or _
       Selection.Range.Style = STYLE_MATCHINGQ_FIXANSWER Then
            Selection.TypeText text:=" " & Chr(11)
            Selection.Paste
    Else 'Error - give instructions:
        MsgBox "Pastes an image from the Clipboard. " & vbCr & _
               "Place the cursor at the end of the question. ", vbExclamation
    End If

End Sub

Public Sub MarkTrueAnswer(control As IRibbonControl)
' Marks the right answer or switches true and false statements.
    If Selection.Range.Style = STYLE_CORRECTANSWER Then
        Selection.Range.Style = STYLE_INCORRECTANSWER
    ElseIf Selection.Range.Style = STYLE_INCORRECTANSWER Then
        Selection.Range.Style = STYLE_CORRECTANSWER
     ElseIf Selection.Range.Style = STYLE_TRUESTATEMENT Then
        Selection.Range.Style = STYLE_FALSESTATEMENT
    ElseIf Selection.Range.Style = STYLE_FALSESTATEMENT Then
        Selection.Range.Style = STYLE_TRUESTATEMENT
       
    Else 'Error: give instructions:
    MsgBox "This command toggles a statement from True to False." & vbCr & _
           "Cursor must be on an answer for Multiple Choice" & vbCr & _
           "or on a True or False statement.", vbExclamation
    End If
End Sub

' From Daniel: Changes Shuffleanswerfalse
Public Sub ChangeShuffleanswerTrueFalse(control As IRibbonControl)
    If Selection.Range.Style = STYLE_MATCHINGQ Then
        Selection.Range.Style = STYLE_MATCHINGQ_FIXANSWER
    ElseIf Selection.Range.Style = STYLE_MATCHINGQ_FIXANSWER Then
        Selection.Range.Style = STYLE_MATCHINGQ
    ElseIf Selection.Range.Style = STYLE_MULTICHOICEQ Then
        Selection.Range.Style = STYLE_MULTICHOICEQ_FIXANSWER
    ElseIf Selection.Range.Style = STYLE_MULTICHOICEQ_FIXANSWER Then
        Selection.Range.Style = STYLE_MULTICHOICEQ
        
    Else 'Error: give instructions:
    MsgBox "This command is only for MCQs and Matching Questions. " & vbCr & _
           "Place the cursor in the text of the question, then push this button." & vbCr & _
       "Blue Text = Answers are fixed, Black Text = Answers are randomly shuffled.", vbExclamation
    End If
End Sub


' Add a new paragraph with a specified style and text
' Inserted text is selected
Private Sub AddParagraphOfStyle(aStyle, text)
    Set myRange = Selection.Range 'ActiveDocument.Content
    With myRange
        .InsertParagraphBefore
        '.Move Unit:=wdParagraph, Count:=1
        .Style = aStyle
        .InsertBefore text
        .Select
    End With
End Sub

' Moodle requires a dot (.) as a decimal separator. Thus, all comma separators need to
' be converted.
Private Sub ConvertDecimalSeparator(ByVal aRange As Range)
    aRange.Find.Execute FindText:=",", ReplaceWith:=".", _
    Format:=False, Replace:=wdReplaceAll
End Sub

' Remove all formatting from the document
Private Sub RemoveFormatting()
    Selection.WholeStory
    Selection.Find.ClearFormatting
End Sub

' Count the number of paragraphs having the specified
' style in the defined range
Function CountStylesInRange(aStyle, startPoint, endPoint) As Integer
   Dim counter
   Set aRange = ActiveDocument.Range(Start:=startPoint, End:=endPoint)
   endP = aRange.End  'store end point
   counter = 0
   
   With aRange.Find
        .ClearFormatting
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Style = aStyle
        .Format = True
        Do While .Execute(Wrap:=wdFindStop) = True
            If aRange.End > endP Then
                Exit Do
            Else
                counter = counter + 1    ' Increment Counter.
            End If
        Loop
        
    End With
    CountStylesInRange = counter
End Function

' Check if the specified style is found in the range
Function StyleFound(aStyle, aRange As Range) As Boolean
   With aRange.Find
        .ClearFormatting
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Style = aStyle
        .Format = True
        .Execute Wrap:=wdFindStop
    End With
    'MsgBox "Style: " & aStyle & " Found: " & aRange.Find.Found
    StyleFound = aRange.Find.Found
End Function

' Removes answer weights from the selection
Public Sub RemoveAnswerWeightsFromTheSelection()
    With Selection.Find
        .ClearFormatting
        .Style = STYLE_ANSWERWEIGHT
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


' Checks the questionnaire.
' Returns true if everyhing is fine, otherwise false
Function CheckQuestionnaire() As Boolean
    If ActiveDocument.Content.Characters.Count = 1 Then Exit Function
  
    Dim startOfQuestion, endOfQuestion, setEndPoint
    isOk = True
    setEndPoint = False ' indicates whether the question end point should be set
    startOfQuestion = 0
    questionType = ""
   
    ' Check each paragraph at a time and specify needed tags
    For Each para In ActiveDocument.Paragraphs
        
        ' Check if empty paragraph
        If para.Range = vbCr Then
            para.Range.Delete ' delete all empty paragraphs
            If questionType = "" Then questionType = para.Range.Style.NameLocal
        ElseIf para.Range.Style.NameLocal = STYLE_MULTICHOICEQ Or _
               para.Range.Style.NameLocal = STYLE_MULTICHOICEQ_FIXANSWER Or _
               para.Range.Style.NameLocal = STYLE_MATCHINGQ Or _
               para.Range.Style.NameLocal = STYLE_MATCHINGQ_FIXANSWER Or _
               para.Range.Style.NameLocal = STYLE_NUMERICALQ Or _
               para.Range.Style.NameLocal = STYLE_SHORTANSWERQ Then
            
            If setEndPoint Then
                endOfQuestion = para.Range.Start
                isOk = CheckQuestion(startOfQuestion, endOfQuestion)
                If isOk = False Then Exit For ' Exit if error is found
            End If
            
            startOfQuestion = para.Range.Start
            setEndPoint = True
            questionType = para.Range.Style.NameLocal
                        
        ElseIf para.Range.Style.NameLocal = STYLE_TRUESTATEMENT Or _
               para.Range.Style.NameLocal = STYLE_FALSESTATEMENT Or _
               para.Range.Style.NameLocal = STYLE_MISSINGWORDQ Or _
               para.Range.Style.NameLocal = STYLE_BLANK_WORD Or _
               para.Range.Style.NameLocal = STYLE_ESSAY Then
               
            If setEndPoint Then
                endOfQuestion = para.Range.Start
                isOk = CheckQuestion(startOfQuestion, endOfQuestion)
                If isOk = False Then Exit For ' Exit if error is found
                startOfQuestion = para.Range.Start
            End If
            
            questionType = para.Range.Style.NameLocal
           
            isOk = CheckQuestion(startOfQuestion, para.Range.End)
            If isOk = False Then Exit For ' Exit if error is found
            startOfQuestion = para.Range.End
            questionType = "NOT_KNOWN"
            setEndPoint = False
        ElseIf para.Range.Style.NameLocal = STYLE_CORRECTANSWER And _
               questionType = STYLE_NUMERICALQ Then
            ' Exit if error is found
            If CheckNumericAnswer(para.Range) = False Then Exit Function
        End If
        
        ' Check if the end of document
        If para.Range.End = ActiveDocument.Range.End And _
        startOfQuestion <> para.Range.End Then
           isOk = CheckQuestion(startOfQuestion, ActiveDocument.Range.End)
        End If
        
        If isOk = False Then Exit For ' Exit if error is found
        
    Next para
    
    If ActiveDocument.Content.Characters.Count = 1 Then Exit Function
    
    Application.ScreenRefresh
    CheckQuestionnaire = isOk
End Function

' Checks whether the chosen question is valid
' Returns true if the question is OK, otherwise
Function CheckQuestion(startPoint, endPoint) As Boolean
    Dim isOk
    Set aRange = ActiveDocument.Range(startPoint, endPoint)
    aRange.Select
    'MsgBox "See Range for specifying question type." & questionType & vbCr & _
    '      "Start: " & startPoint & " End: " & endPoint
    isOk = True 'no errors
        
    If questionType = STYLE_MULTICHOICEQ Or _
      questionType = STYLE_MULTICHOICEQ_FIXANSWER Then
    
        ' Check that there are right anwers specified
        rightCount = CountStylesInRange(STYLE_CORRECTANSWER, startPoint, endPoint)
                      
        If rightCount = 0 Then
            aRange.Select
            MsgBox "Error, no correct answer defined.", vbExclamation
            isOk = False
        End If
        
    ElseIf questionType = STYLE_SHORTANSWERQ Then
        rightCount = CountStylesInRange(STYLE_SHORT_ANSWER, startPoint, endPoint)
                      
        If rightCount = 0 Then
            aRange.Select
            MsgBox "Error, no correct short answer is defined.", vbExclamation
            isOk = False
        End If
        
    ElseIf questionType = STYLE_NUMERICALQ Then
        rightCount = CountStylesInRange(STYLE_SHORT_ANSWER, startPoint, endPoint)
                  
        If rightCount = 0 Then
            aRange.Select
            MsgBox "Error, no correct numerical answer is defined.", vbExclamation
            isOk = False
        End If
    
    ' MATCHING QUESTION
    ElseIf questionType = STYLE_MATCHINGQ Or questionType = STYLE_MATCHINGQ_FIXANSWER Then
        
       ' Count the number of pairs
       rightPairCount = CountStylesInRange(STYLE_RIGHT_PAIR, startPoint, endPoint)
       leftPairCount = CountStylesInRange(STYLE_LEFT_PAIR, startPoint, endPoint)
       
       ' Too few pairs
       If leftPairCount < 3 Then
           aRange.Select
           MsgBox "Error, there are not enough pairs for a matching question" & vbCr & _
                  "There must be at least 3 matching pairs. Please add more.", vbExclamation, "Error!"
           isOk = False
       ' Error -> the number of left and right pairs is different or zero
       ElseIf rightPairCount <> leftPairCount Then
           aRange.Select
           MsgBox "Error, pairs are not correctly defined" & vbCr & _
                  "The number of left and right pairs is not equal.", vbExclamation, "Error!"
           isOk = False
       End If
       
    ElseIf questionType = STYLE_MISSINGWORDQ Then
        wordCount = CountStylesInRange(STYLE_BLANK_WORD, startPoint, endPoint)
        If wordCount <> 1 Then
            aRange.Select
            MsgBox "There must be exactly one answer specified as a blank word." _
            + Chr(13) + Chr(13) + "To remove unnecessary markup, select a word(s) and press Ctrl+Space.", vbExclamation, "Error!"
            isOk = False
        End If
            
    ElseIf questionType = STYLE_TRUESTATEMENT Or _
           questionType = STYLE_FALSESTATEMENT Or _
           questionType = STYLE_ESSAY Then
           'nothing to check for answers to these ones. Figure out what the issue is with being last question in test and fix here?
       
    ' UNDEFINED QUESTION TYPE
    Else
        aRange.Select
        MsgBox "Undefined Question type:" & questionType & vbCr & vbCr _
               & "Illegal question is deleted.", vbExclamation, "Error!"
        aRange.Delete
    End If
    
    'MsgBox questionType & "=" & isOk  'debugging output - show OK after each q checked.
    CheckQuestion = isOk
End Function

' Check the numeric answer. Note, not checking all valid GIFT formats.
Function CheckNumericAnswer(aRange As Range) As Boolean
    
    CheckNumericAnswer = True ' By default OK
    
    ' Search for the error margin separator
    aRange.Find.Execute FindText:=":", Format:=False
    
    If aRange.Find.Found = False And IsNumeric(aRange) = False Then
            aRange.Select
            Response = MsgBox("Is this a right numerical answer?" & vbCr & _
                       "Your answer: " & aRange, vbYesNo, "Correct?")
            If Response = vbNo Then CheckNumericAnswer = False
    End If
End Function

' Inserts text before the specified range. A new paragraph is inserted.
Sub InsertTextBeforeRange(ByVal text As String, ByVal aRange As Range)
    With aRange
         .Style = STYLE_NORMAL
        .InsertBefore text
        .Move Unit:=wdParagraph, Count:=1
    End With

End Sub

' Inserts text having trailing VbCr before the range
Sub InsertQuestionEndTag(endPoint As Integer)
    Set aRange = ActiveDocument.Range(endPoint - 1, endPoint)
    With aRange
        .InsertBefore vbCr & TAG_QUESTION_END
        .Style = STYLE_NORMAL
        .Move Unit:=wdParagraph, Count:=1
    End With

End Sub

' Inserts text at the end of the paragraph before the trailing VbCr
Sub InsertAfterBeforeCR(ByVal text As String, ByVal aRange As Range)
    aRange.End = aRange.End - 1 ' insert text before cr
    With aRange
        .InsertAfter text
        .Style = STYLE_NORMAL
        .Move Unit:=wdParagraph, Count:=1
    End With
End Sub

' Inserts text at the end of the paragraph
Sub InsertAfterRange(ByVal text As String, aStyle, ByVal aRange As Range)
    With aRange
        .EndOf Unit:=wdParagraph, Extend:=wdMove
        .InsertParagraphBefore
        .Move Unit:=wdParagraph, Count:=-1
        .Style = aStyle
        .InsertBefore text
        .Select
    End With
End Sub

' Set the answer weights of multiple choice questions.
Public Sub SetAnswerWeights() ' aStyle, startPoint, endPoint)
    Dim startPoint, endPoint, rightScore, wrongScore
    
    If Selection.Range.Style = STYLE_MULTICHOICEQ Or STYLE_MULTICHOICE_FIXANSWER Then
        startPoint = Selection.Paragraphs(1).Range.Start
        rightCount = 0
        wrongCount = 0
        Selection.MoveDown Unit:=wdParagraph, Count:=1
        
        Do While Selection.Range.Style = STYLE_CORRECTANSWER Or _
              Selection.Range.Style = STYLE_INCORRECTANSWER Or _
              Selection.Range.Style = STYLE_FEEDBACK Or _
              Selection.Range.Style = STYLE_ANSWERWEIGHT
                          
            'Delete empty paragraphs
            If Selection.Paragraphs(1).Range = vbCr Then
                Selection.Paragraphs(1).Range.Delete ' delete all empty paragraphs
            ' Remove old answer weights
            ElseIf Selection.Range.Style = STYLE_ANSWERWEIGHT Then
                With Selection.Find
                    .ClearFormatting
                    .Style = STYLE_ANSWERWEIGHT
                    .text = ""
                    .Replacement.text = ""
                    .Forward = True
                    .Format = True
                    .Execute Replace:=wdReplaceOne
                End With
            End If
            
            ' Count the number of right and wrong answers
            If Selection.Range.Style = STYLE_CORRECTANSWER Then
                rightCount = rightCount + 1
            ElseIf Selection.Range.Style = STYLE_INCORRECTANSWER Then
                wrongCount = wrongCount + 1
            End If
            
           If Selection.Paragraphs(1).Range.End = ActiveDocument.Range.End Then
                endPoint = Selection.Paragraphs(1).Range.End
                Exit Do
           Else
                Selection.MoveDown Unit:=wdParagraph, Count:=1
                endPoint = Selection.Paragraphs(1).Range.Start
           End If
        Loop
        
        Set QuestionRange = ActiveDocument.Range(startPoint, endPoint)
        
        If rightCount < 1 Then
            QuestionRange.Select
            MsgBox "No correct answer specified.", vbExclamation, "Error!"
        Else
            ' Calculate the right and wrong scores
            rightScore = Round(100 / rightCount, 3)
            ' MODIFY the default scoring principle for wrong answers if necessary
            wrongScore = -rightScore
            
            AddAnswerWeights QuestionRange, rightScore, wrongScore
        End If
    Else
        MsgBox "Place the cursor on the question title" & vbCr & _
               "of the Multiple Choice Question", vbExclamation, "Error!"
        ' Find the previous paragraph having the style of multiple choice question.
        With Selection.Find
            .ClearFormatting
            .text = ""
            .Style = STYLE_MULTICHOICEQ Or STYLE_MULTICHOICEQ_FIXANSWER
            .Forward = False
            .Format = True
            .MatchCase = False
            .Execute
        End With
    End If
End Sub

' Insert answer weights
Private Sub AddAnswerWeights(ByVal aRange As Range, rightScore, wrongScore)
    ' Check each paragraph at a time and specify needed tags
    For Each para In aRange.Paragraphs
        ' Check if empty paragraph
        If para.Range = vbCr Then
            para.Range.Delete ' delete all empty paragraphs
        ElseIf para.Range.Style = STYLE_CORRECTANSWER Then
            InsertAnswerWeight rightScore, para.Range
        ElseIf para.Range.Style = STYLE_INCORRECTANSWER Then
            InsertAnswerWeight wrongScore, para.Range
        End If
    Next para

End Sub

' Inserts text at the end of the chapter before the trailing VbCr
Private Sub InsertAnswerWeight(Score, ByVal aRange As Range)
    startPoint = aRange.Start
    scoreString = "" & Score & "%"
    aRange.InsertBefore scoreString
    Set newRange = ActiveDocument.Range(Start:=startPoint, End:=startPoint + Len(scoreString))
    newRange.Style = STYLE_ANSWERWEIGHT
    'Moodle requires that the decimal separator is dot, not comma.
    ConvertDecimalSeparator newRange
    
End Sub

'LTG: This seems to be unused? left over from GIFT format days?
Private Sub FindBlanks(aRange As Range)
    'Set aRange = Selection.Paragraphs(1).Range
    
    
    'ActiveDocument.Content
    endPoint = aRange.End
    With aRange.Find
        .ClearFormatting
        .Style = STYLE_BLANK_WORD
        Do While .Execute(FindText:="", Forward:=True, Format:=True) = True And _
            Selection.Range.End < endPoint And ActiveDocument.Range.End <> Selection.Range.End
            With .Parent
                .InsertBefore "{="
                    .InsertAfter "}"
              '  .Move Unit:=wdParagraph, Count:=1
              .Move Unit:=wdWord, Count:=1
            End With
        Loop
    End With
End Sub


Public Sub Export(control As IRibbonControl)

    StatusBar = "Checking the quiz questions formatting, please wait..."
    ' Before conversion, document is validated
    If CheckQuestionnaire = True Then
    
        StatusBar = "Converting to Moodle XML format, please wait..."
        Convert2XML
        
    Else
        MsgBox "The export operation can not be started until everything is OK" & vbCr & "and there is at least one question.", vbCritical, "Error"
    End If


End Sub

Private Sub Convert2XML()


' Macro recorded on 21.12.2008 by Daniel Refresh Header (translation?)
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.Fields.Update
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Fields.Update
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    
    'look for the folder containing .xml question patterns
    xmlpath = ActiveDocument.Path & "\xml-question\"
    If Not DirExists(xmlpath) Then
        If ActiveDocument.AttachedTemplate.Path <> "" Then
            xmlpath = ActiveDocument.AttachedTemplate.Path & "\xml-question\"
        Else
            xmlpath = ActiveDocument.Path & "\xml-question\"
        End If
        If Not DirExists(xmlpath) Then
             MsgBox "The xml-question\ folder is not found. Please keep this quiz question document in the original folder.", vbCritical, "Error"
             Exit Sub
        End If
    End If
   
    'choose the file name to save with
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    'fd.FilterIndex = 2 fuer Word 2003, 14 fuer Word 2010
    fd.FilterIndex = 14
    fd.InitialFileName = FILE_PREFIX & Format(Date, "yyyymmdd") & ".xml"
    
    If fd.Show <> -1 Then Exit Sub

     
    Dim header As String
    header = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.text
    

     '//*** save the file in utf-8 using stream ***//
    Dim objStream 'As ADODB.Stream ' CPF took off the type
    'Create the stream
     Set objStream = CreateObject("ADODB.Stream")
    'Initialize the stream
       objStream.Open
     'Reset the position and indicate the charactor encoding
    objStream.Position = 0
    objStream.Charset = "UTF-8"
    
    'specify XML version and that it is a quiz
    objStream.WriteText "<?xml version=""1.0""?><quiz>" & vbCr
    'write the categories from the header
    objStream.WriteText "<question type=""category"">" & vbCr
    objStream.WriteText "<category>" & vbCr
    objStream.WriteText "<text>" & header & "</text>" & vbCr
    objStream.WriteText "</category>" & vbCr
    objStream.WriteText "</question>" & vbCr & vbCr
    
    Dim dd As MSXML2.DOMDocument60
    Dim xmlnod As IXMLDOMNode
    Dim xmlnodelist As IXMLDOMNodeList
    Dim para As Paragraph, paralookahead As Paragraph
  
    For Each para In ActiveDocument.Paragraphs '?handle each paragraph separately.
        Set dd = New MSXML2.DOMDocument60
        
        Select Case para.Style
            
        Case STYLE_SHORTANSWERQ
                dd.Load xmlpath & "shortanswer.xml"
                ProcessCommonTags dd, para
                ' processing each <answer>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("answer")
                dd.documentElement.RemoveChild xmlnod
                Do While (paralookahead.Style = STYLE_SHORT_ANSWER)
                    xmlnod.Attributes.getNamedItem("fraction").text = "100"
                    xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                    dd.documentElement.appendChild xmlnod
                    
                    Set xmlnod = xmlnod.CloneNode(True)
                    Set paralookahead = paralookahead.Next
                    If paralookahead Is Nothing Then Exit Do
                Loop
                
             Case STYLE_ESSAY
                dd.Load xmlpath & "essay.xml"
                ProcessCommonTags dd, para
                'LTG: do I need to have Set paralookahead = para.Next here? why / why not?
            
            Case STYLE_NUMERICALQ
                dd.Load xmlpath & "numerical.xml"
                ProcessCommonTags dd, para
                ' processing each <answer>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("answer")
                dd.documentElement.RemoveChild xmlnod
                
                Do While (paralookahead.Style = STYLE_SHORT_ANSWER)
                   xmlnod.Attributes.getNamedItem("fraction").text = "100"
                   xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                   Set paralookahead = paralookahead.Next
                   If Not paralookahead Is Nothing Then
                     If (paralookahead.Style = STYLE_NUM_TOLERANCE) Then
                       xmlnod.SelectSingleNode("tolerance").text = RemoveCR(paralookahead.Range.text)
                       Set paralookahead = paralookahead.Next
                     Else
                       xmlnod.SelectSingleNode("tolerance").text = "0"
                     End If
                   End If
                   dd.documentElement.appendChild xmlnod
                   Set xmlnod = xmlnod.CloneNode(True)
                   If paralookahead Is Nothing Then Exit Do
                Loop
             
            Case STYLE_FALSESTATEMENT
                dd.Load xmlpath & "false.xml"
                ProcessCommonTags dd, para
                Set paralookahead = para.Next
                                
            Case STYLE_TRUESTATEMENT
                dd.Load xmlpath & "true.xml"
                ProcessCommonTags dd, para
                Set paralookahead = para.Next
    
             Case STYLE_MULTICHOICEQ_FIXANSWER
                dd.Load xmlpath & "multichoicefix.xml"
                ProcessCommonTags dd, para
                
                ' processing each <answer>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("answer")
                dd.documentElement.RemoveChild xmlnod
                rac = 0
                wac = 0
                Do While (paralookahead.Style = STYLE_CORRECTANSWER) Or (paralookahead.Style = STYLE_INCORRECTANSWER)
                    If paralookahead.Style = STYLE_CORRECTANSWER Then
                        xmlnod.Attributes.getNamedItem("fraction").text = "100"
                        rac = rac + 1
                    Else
                        xmlnod.Attributes.getNamedItem("fraction").text = "0"
                        wac = wac + 1
                    End If
                    xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                    dd.documentElement.appendChild xmlnod
                    
                    Set xmlnod = xmlnod.CloneNode(True)
                    Set paralookahead = paralookahead.Next
                    If paralookahead Is Nothing Then Exit Do
                Loop
                
                If rac > 1 Then
                    ' multiple correct/incorrect answers
                    dd.documentElement.SelectSingleNode("single").text = "false"
                    ' re-looping for setting multi-true-answer fractions
                    For Each mansw In dd.documentElement.SelectNodes("answer")
                        With mansw.Attributes.getNamedItem("fraction")
                            If .text = 100 Then .text = Replace(100 / rac, ",", ".")
                            If .text = 0 Then .text = Replace(-100 / wac, ",", ".")
                        End With
                    Next mansw
                End If
                            
             Case STYLE_MULTICHOICEQ
                dd.Load xmlpath & "multichoicevar.xml"
                ProcessCommonTags dd, para
                
                ' processing each <answer>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("answer")
                dd.documentElement.RemoveChild xmlnod
                rac = 0 'right answer choices
                wac = 0 'wrong answer choices
                Do While (paralookahead.Style = STYLE_CORRECTANSWER) Or (paralookahead.Style = STYLE_INCORRECTANSWER)
                    If paralookahead.Style = STYLE_CORRECTANSWER Then
                        xmlnod.Attributes.getNamedItem("fraction").text = "100"
                        rac = rac + 1
                    Else
                        xmlnod.Attributes.getNamedItem("fraction").text = "0"
                        wac = wac + 1
                    End If
                    xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                    dd.documentElement.appendChild xmlnod
                    
                    Set xmlnod = xmlnod.CloneNode(True)
                    Set paralookahead = paralookahead.Next
                    ' Feedback Style processing here
                    If paralookahead.Style = STYLE_FEEDBACK Then
                        ' Set XML <feedback> text
                        xmlnod.SelectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.text)
                        Set paralookahead = paralookahead.Next
                    End If
                    
                    dd.documentElement.appendChild xmlnod
                    Set xmlnod = xmlnod.CloneNode(True)
                                        
                    If paralookahead Is Nothing Then Exit Do
                Loop
                
                If rac > 1 Then
                    ' multiple correct/incorrect answers
                    dd.documentElement.SelectSingleNode("single").text = "false"
                    ' re-looping for setting multi-true-answer fractions
                    For Each mansw In dd.documentElement.SelectNodes("answer")
                        With mansw.Attributes.getNamedItem("fraction")
                            If .text = 100 Then .text = Replace(100 / rac, ",", ".") 'original
                            If .text = 0 Then .text = Replace(-100 / wac, ",", ".") 'original
                            'If .text = 100 Then .text = Round(100 / rac, 5) 'sd 2010 für Moodle 2.0
                            'If .text = 0 Then .text = Round(-100 / wac, 5) 'sd 2010 für moodle 2.0
                        End With
                    Next mansw
                End If
                            
            Case STYLE_MATCHINGQ
                dd.Load xmlpath & "matchingvar.xml"
                ProcessCommonTags dd, para
                
                ' processing each <subquestion>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("subquestion")
                dd.documentElement.RemoveChild xmlnod
                Do While (paralookahead.Style = STYLE_LEFT_PAIR) Or (paralookahead.Style = STYLE_RIGHT_PAIR)
                    If paralookahead.Style = STYLE_LEFT_PAIR Then
                        xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                    Else
                        xmlnod.SelectSingleNode("answer").SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                        
                        dd.documentElement.appendChild xmlnod
                        Set xmlnod = xmlnod.CloneNode(True)
                    End If
                    Set paralookahead = paralookahead.Next
                    If paralookahead Is Nothing Then Exit Do
                Loop
            
            Case STYLE_MATCHINGQ_FIXANSWER
                dd.Load xmlpath & "matchingfix.xml"
                ProcessCommonTags dd, para
                
                ' processing each <subquestion>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("subquestion")
                dd.documentElement.RemoveChild xmlnod
                Do While (paralookahead.Style = STYLE_LEFT_PAIR) Or (paralookahead.Style = STYLE_RIGHT_PAIR)
                    If paralookahead.Style = STYLE_LEFT_PAIR Then
                        xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                    Else
                        xmlnod.SelectSingleNode("answer").SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                        
                        dd.documentElement.appendChild xmlnod
                        Set xmlnod = xmlnod.CloneNode(True)
                    End If
                    Set paralookahead = paralookahead.Next
                    If paralookahead Is Nothing Then Exit Do
                Loop
            
            Case STYLE_MISSINGWORDQ
                dd.Load xmlpath & "shortanswer.xml"
                Dim char As Range
                misword = ""
                For Each char In para.Range.Characters
                    If char.Style = STYLE_BLANK_WORD Then misword = misword & char.text
                Next char
                ProcessCommonTags dd, para
                dd.documentElement.SelectSingleNode("name").SelectSingleNode("text").text = Replace( _
                  dd.documentElement.SelectSingleNode("name").SelectSingleNode("text").text, misword, "__________")
                
                ' processing each <answer>'
                Set paralookahead = para.Next
                Set xmlnod = dd.documentElement.SelectSingleNode("answer")
                xmlnod.Attributes.getNamedItem("fraction").text = "100"
                xmlnod.SelectSingleNode("text").text = misword
                            
            Case STYLE_COMMENT
                Comment = "<!-- " & RemoveCR(para.Range.text) & " -->"
                objStream.WriteText (Comment & vbCr & vbCr)
                dd.loadXML ("")
            
            Case Else
                dd.loadXML ("")
        End Select
        
        
          If Not paralookahead Is Nothing Then
          If (paralookahead.Style = STYLE_QUESTIONNAME) Then
             Set xmlnod = dd.documentElement.SelectSingleNode("name")
             dd.documentElement.RemoveChild xmlnod
             xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
             dd.documentElement.appendChild xmlnod
             Set xmlnod = xmlnod.CloneNode(True)
             Set paralookahead = paralookahead.Next
          End If
           
        End If
                If Not paralookahead Is Nothing Then '**seems to be setting generalfeedback for any feedback tag...
'          If (paralookahead.Style = STYLE_FEEDBACK) Then
'             Set xmlnod = dd.documentElement.SelectSingleNode("generalfeedback")
'             dd.documentElement.RemoveChild xmlnod
'             xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
'             dd.documentElement.appendChild xmlnod
'             Set xmlnod = xmlnod.CloneNode(True)
'             Set paralookahead = paralookahead.Next
'          End If
        End If

        If dd.XML <> "" Then objStream.WriteText (dd.XML & vbCr)
        
        Set dd = Nothing
    Next para
    
 objStream.WriteText "</quiz>"
 'Save the stream to a file
 objStream.SaveToFile FileName:=fd.SelectedItems(1), Options:=adSaveCreateOverWrite
     
End Sub

'This is called as the first processing task for each question. It
Private Sub ProcessCommonTags(dd As DOMDocument60, para As Paragraph)
                ' processing <name> '
                dd.documentElement.SelectSingleNode("name") _
                .SelectSingleNode("text").text = RemoveCR(para.Range.text)
                
                ' processing <questiontext> '
                dd.documentElement.SelectSingleNode("questiontext") _
                .SelectSingleNode("text").text = XSLT_Range(para.Range, "FormattedText.xslt")
                
                '
                If Not XSLT_Range(para.Range, "PictureName.xslt") = "" Then 'if it is NOT null/empty
                    
                        Dim header As String
                        Dim stringlength As Long
                        header = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.text
                        stringlength = Len(header)
                        header = Left(header, stringlength - 1)
                    
                    'processing <image>'
                    dd.documentElement.SelectSingleNode("image").text = "Images_forQuizQuestions/" & header & Right(XSLT_Range(para.Range, "PictureName.xslt"), 4)
                    'dd.documentElement.SelectSingleNode("image").text = Mid(XSLT_Range(para.Range, "PictureName.xslt"), 10) (commented out: Rohrer)'
                    'processing <image_base64>'
                    dd.documentElement.SelectSingleNode("image_base64").text = XSLT_Range(para.Range, "Picture.xslt")
                End If
                
                
                

End Sub


Private Function XSLT_Range(textrange As Range, xsltfilename As String) As String
    Dim xsldoc As New MSXML2.FreeThreadedDOMDocument60
    xsldoc.Load (xmlpath & xsltfilename)
    Dim xslt As New MSXML2.XSLTemplate60
    Set xslt.StyleSheet = xsldoc
    Dim xsltProcessor As IXSLProcessor
    Set xsltProcessor = xslt.createProcessor
    Dim d As New MSXML2.DOMDocument60
    d.loadXML textrange.XML '!!! Bug in Word 2010 when file is created from a template (Textrange.xml can not be read)
    xsltProcessor.input = d
    xsltProcessor.transform
    s = xsltProcessor.output
    
    Set xsltProcessor = Nothing
    Set xslt = Nothing
    Set xsldoc = Nothing
    XSLT_Range = s
End Function


Private Function RemoveCR(str As String) As String
    str = Replace(str, vbCr, "")
    RemoveCR = Trim(Application.CleanString(str))
End Function


Function DirExists(ByVal sDirName As String) As Boolean
On Error Resume Next
DirExists = (GetAttr(sDirName) And vbDirectory) = vbDirectory
Err.Clear
End Function


Public Sub Check(control As IRibbonControl)
' Macro recorded on 21.12.2008 by Daniel to Update Header
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.Fields.Update
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Fields.Update
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    If CheckQuestionnaire Then MsgBox "Now everything is OK", vbInformation
End Sub




