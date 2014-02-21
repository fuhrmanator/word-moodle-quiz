'TODO:  Follow these steps to enable the Ribbon (XML) item:

'1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New MoodleQuestions()
'End Function

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.
'FIXED (didn't do, different design, simplify) integrate template file so Paragraph Styles will be valid
'FIXED add all Callbacks from ribbon buttons 
'FIXED add icons to Ribbon item
'FIXED fix paragraph styles so language comes from Keyboard or Normal (http://answers.microsoft.com/en-us/office/forum/office_2010-word/how-to-specify-dont-change-the-language-setting-in/966aec6e-4d4d-4fef-af42-5c4ad260f751)
'TODO find a deployment site for Project Publishing. Google Drive won't work because it doesn't have clean URLs for directories.
'TODO Try using style feedback to indicate [shuffled] questions (rather than arbitrary colors). Numbering allows inserting text after the 1. (e.g., 1. [shuffled])
'TODO Fix feedback button (for which questions is answer feedback valid? Do we need different feedback word styles?)



Imports Microsoft.Office.Interop.Word
Imports stdole

<Runtime.InteropServices.ComVisible(True)> _
Public Class MoodleQuestions
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("MoodleQuestions.MoodleQuestions.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        Me.ribbon.ActivateTab("MoodleQuestions") 'Make Moodle Questions toolbar active on startup
    End Sub

    Public Function OnLoadImage(imageId As String) As IPictureDisp
        Dim tempImage As stdole.IPictureDisp = Nothing
        'load image from resources file
        tempImage = Microsoft.VisualBasic.Compatibility.VB6.Support.ImageToIPicture(My.Resources.RibbonIcons.ResourceManager.GetObject(imageId))
        Return tempImage
    End Function
    '''''BUTTON callbacks

    ' Add Multiple Choice Question to the end of the active document
    Public Sub AddMultipleChoiceQ(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_MULTICHOICEQ, "Insert Multiple Choice Question")
    End Sub

    ' Add Matching Question to the end of the active document
    Public Sub AddMatchingQ(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_MATCHINGQ, "Insert Matching Question")
    End Sub

    ' Add Numerical Question to the end of the active document
    Public Sub AddNumericalQ(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_NUMERICALQ, "Insert Numerical Question")
    End Sub


    ' Add Short Answer Question to the end of the active document
    Public Sub AddShortAnswerQ(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_SHORTANSWERQ, "Insert Short Answer Question")
    End Sub

    ' Add Missing Word Question
    Public Sub AddMissingWordQ(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_MISSINGWORDQ, "Insert Missing Word Question. Then select the missing word!")
    End Sub

    ' Add an Essay
    Public Sub AddEssay(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_ESSAY, "Insert An Essay question here (an Open Question). [This can not be the last question in the document.]")
    End Sub
    Public Sub ToggleMissingWord(ByVal control As Office.IRibbonControl)
        ' Only applies to questions of STYLE_MISSINGWORDQ
        If (getSelectionStyle() = STYLE_MISSINGWORDQ) Then
            ' get only the first word of the selection
            Dim aRange As Microsoft.Office.Interop.Word.Range = getSelectionRange()
            aRange.Start = Globals.ThisDocument.Application.Selection.Words(1).Start
            aRange.End = Globals.ThisDocument.Application.Selection.Words(1).End
            ' toggle the style of the word
            If CType(aRange.Words(1).Style, Word.Style).NameLocal = STYLE_BLANK_WORD Then
                Globals.ThisDocument.Application.Selection.ClearCharacterStyle()
            Else
                aRange.Style = STYLE_BLANK_WORD
            End If
        Else
            MsgBox("Select a word inside a Missing-word question first.", vbExclamation)
        End If


    End Sub

    Public Sub AddQuestionFeedback(ByVal control As Office.IRibbonControl)
        If getSelectionStyle() = STYLE_ANSWERWEIGHT Or _
           getSelectionStyle() = STYLE_SHORTANSWERQ Or _
           getSelectionStyle() = STYLE_MISSINGWORDQ Or _
           getSelectionStyle() = STYLE_CORRECTANSWER Or _
           getSelectionStyle() = STYLE_INCORRECTANSWER Or _
           getSelectionStyle() = STYLE_SHORT_ANSWER Or _
           getSelectionStyle() = STYLE_RIGHT_PAIR Or _
           getSelectionStyle() = STYLE_NUM_TOLERANCE Or _
           getSelectionStyle() = STYLE_TRUESTATEMENT Or _
           getSelectionStyle() = STYLE_FALSESTATEMENT Or _
           getSelectionStyle() = STYLE_QUESTIONNAME Or _
           getSelectionStyle() = STYLE_BLANK_WORD Then
            InsertAfterRange("Insert feedback of the previous choice or answer here.", _
                             STYLE_FEEDBACK, Globals.ThisDocument.Application.Selection.Paragraphs(1).Range)


        Else 'Error: Give Instructions:
            MsgBox("Feedback is placed at the end of the last possible response. " & vbCr & _
                   "It doesn't work for True/False questions." & vbCr & _
                   "Place the cursor on top of the question or answer you are giving feedback for.", vbExclamation)

        End If
    End Sub
    ' Add tolerance
    Public Sub AddNumericalTolerance(ByVal control As Office.IRibbonControl)
        If getSelectionStyle() = STYLE_SHORT_ANSWER Then
            InsertAfterRange("Replace me with Tolerance for the answer as a Decimal. Eg: 0.01", _
                             STYLE_NUM_TOLERANCE, Globals.ThisDocument.Application.Selection.Paragraphs(1).Range)
        Else 'Error: Give Instructions:
            MsgBox(" " & vbCr & _
                   "Place the cursor at the end of the numerical answer.", vbExclamation)
        End If
    End Sub

    ' Add QuestionName / Question Title
    Public Sub AddQuestionTitle(ByVal control As Office.IRibbonControl)
        If getSelectionStyle() = STYLE_ANSWERWEIGHT Or _
           getSelectionStyle() = STYLE_SHORTANSWERQ Or _
           getSelectionStyle() = STYLE_MISSINGWORDQ Or _
           getSelectionStyle() = STYLE_CORRECTANSWER Or _
           getSelectionStyle() = STYLE_NUM_TOLERANCE Or _
           getSelectionStyle() = STYLE_INCORRECTANSWER Or _
           getSelectionStyle() = STYLE_TRUESTATEMENT Or _
           getSelectionStyle() = STYLE_SHORT_ANSWER Or _
           getSelectionStyle() = STYLE_FALSESTATEMENT Or _
           getSelectionStyle() = STYLE_RIGHT_PAIR Or _
           getSelectionStyle() = STYLE_BLANK_WORD Then
            InsertAfterRange("Add a question title.", _
                 STYLE_QUESTIONNAME, Globals.ThisDocument.Application.Selection.Paragraphs(1).Range)
        Else 'Error: Give Instructions:
            MsgBox("Feedback to insert at the end of the last response selected. " & vbCr & _
                   "The title must appear before the feedback" & vbCr & _
                   "Place the cursor at the end of the last line selected", vbExclamation)
        End If
    End Sub


    ' Add a true statement of the true-false question
    Public Sub AddTrueStatement(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_TRUESTATEMENT, "True-false question: insert a TRUE statement here (not at the end of the document)")
    End Sub

    ' Add a false statement of the true-false question
    Public Sub AddFalseStatement(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_FALSESTATEMENT, "True-false question: insert a FALSE statement here (not at the end of the document)")
    End Sub

    ' Add a comment
    Public Sub AddComment(ByVal control As Office.IRibbonControl)
        AddParagraphOfStyle(STYLE_COMMENT, "")
    End Sub
    Public Sub PasteImage(ByVal control As Office.IRibbonControl)
        '  Adds an image from the clipboard into a question.
        With Globals.ThisDocument.Application.Selection
            If .Range.Style = STYLE_SHORTANSWERQ Or _
               .Range.Style = STYLE_MISSINGWORDQ Or _
               .Range.Style = STYLE_MULTICHOICEQ Or _
               .Range.Style = STYLE_MATCHINGQ Or _
               .Range.Style = STYLE_NUMERICALQ Or _
               .Range.Style = STYLE_TRUESTATEMENT Or _
               .Range.Style = STYLE_FALSESTATEMENT Or _
               .Range.Style = STYLE_MULTICHOICEQ_FIXANSWER Or _
               .Range.Style = STYLE_MATCHINGQ_FIXANSWER Then
                Globals.ThisDocument.Application.Options.ReplaceSelection = False
                .TypeText(Text:=(" " & Chr(11)))
                .Paste()
            Else 'Error - give instructions:
                MsgBox("Pastes an image from the Clipboard. " & vbCr & _
                       "Place the cursor at the end of the question. ", vbExclamation)
            End If
        End With
    End Sub

    Public Sub ToggleAnswer(ByVal control As Office.IRibbonControl)
        'Toggles MCQ answer (right-wrong) or switches true and false statements.
        Dim theStyle As String = getSelectionStyle()

        If theStyle = STYLE_CORRECTANSWER Then
            setSelectionStyle(STYLE_INCORRECTANSWER)
        ElseIf theStyle = STYLE_INCORRECTANSWER Then
            setSelectionStyle(STYLE_CORRECTANSWER)
        ElseIf theStyle = STYLE_TRUESTATEMENT Then
            setSelectionStyle(STYLE_FALSESTATEMENT)
        ElseIf theStyle = STYLE_FALSESTATEMENT Then
            setSelectionStyle(STYLE_TRUESTATEMENT)
        Else 'Error: give instructions:
            MsgBox("This command toggles a statement from True to False." & vbCr & _
                   "Cursor must be on an answer for Multiple Choice" & vbCr & _
                   "or on a True or False statement.", vbExclamation)
        End If
    End Sub

    Public Sub ChangeShuffleanswerTrueFalse(ByVal control As Office.IRibbonControl)

        If getSelectionStyle() = STYLE_MATCHINGQ Then
            setSelectionStyle(STYLE_MATCHINGQ_FIXANSWER)
        ElseIf getSelectionStyle() = STYLE_MATCHINGQ_FIXANSWER Then
            setSelectionStyle(STYLE_MATCHINGQ)
        ElseIf getSelectionStyle() = STYLE_MULTICHOICEQ Then
            setSelectionStyle(STYLE_MULTICHOICEQ_FIXANSWER)
        ElseIf getSelectionStyle() = STYLE_MULTICHOICEQ_FIXANSWER Then
            setSelectionStyle(STYLE_MULTICHOICEQ)

        Else 'Error: give instructions:
            MsgBox("This command is only for MCQs and Matching Questions. " & vbCr & _
                   "Place the cursor in the text of the question, then push this button." & vbCr & _
               "Blue Text = Answers are fixed, Black Text = Answers are randomly shuffled.", vbExclamation)
        End If
    End Sub

    Public Sub Check(ByVal control As Office.IRibbonControl)
        ' Macro recorded on 21.12.2008 by Daniel to Update Header
        Globals.ThisDocument.Application.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
        Globals.ThisDocument.Application.Selection.Fields.Update()
        Globals.ThisDocument.Application.Selection.EndKey(Unit:=WdUnits.wdLine)
        Globals.ThisDocument.Application.Selection.MoveLeft(Unit:=WdUnits.wdCharacter, Count:=1)
        Globals.ThisDocument.Application.Selection.Fields.Update()
        Globals.ThisDocument.Application.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument

        If CheckQuestionnaire() Then MsgBox("Now everything is OK", vbInformation)
    End Sub

    Public Sub Export(ByVal control As Office.IRibbonControl)
        Dim StatusBar As String
        StatusBar = "Checking the quiz questions formatting, please wait..."
        ' Before conversion, document is validated
        If CheckQuestionnaire() = True Then

            StatusBar = "Converting to Moodle XML format, please wait..."
            Convert2XML()

        Else
            MsgBox("The export operation can not be started until everything is OK" & vbCr & "and there is at least one question.", vbCritical, "Error")
        End If


    End Sub

#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region


    ' General purpose styles.
    Public Const STYLE_NORMAL = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal

    Public Const STYLE_FEEDBACK = "A Feedback"
    Public Const STYLE_ANSWERWEIGHT = "A Weight"

    Public Const STYLE_SHORTANSWERQ = "Q Short Answer"
    Public Const STYLE_MULTICHOICEQ = "Q Multi Choice"
    Public Const STYLE_MATCHINGQ = "Q Matching"
    Public Const STYLE_NUMERICALQ = "Q Numerical"
    Public Const STYLE_MISSINGWORDQ = "Q Missing Word"
    Public Const STYLE_TRUESTATEMENT = "Q True Statement"
    Public Const STYLE_FALSESTATEMENT = "Q False Statement"
    Public Const STYLE_CORRECTANSWER = "A Correct Choice"
    Public Const STYLE_INCORRECTANSWER = "A Incorrect Choice"
    Public Const STYLE_SHORT_ANSWER = "A Short Answer"
    Public Const STYLE_LEFT_PAIR = "A Matching Left"
    Public Const STYLE_RIGHT_PAIR = "A Matching Right"
    Public Const STYLE_BLANK_WORD = "MissingWord"
    Public Const STYLE_COMMENT = "Comment"
    ' Supplement(ed by) Daniel
    Public Const STYLE_MULTICHOICEQ_FIXANSWER = "Q Multi Choice FixAnswer"
    Public Const STYLE_MATCHINGQ_FIXANSWER = "Q Matching FixAnswer"
    Public Const STYLE_NUM_TOLERANCE = "Num Tolerance"
    Public Const STYLE_QUESTIONNAME = "Questionname"
    'from v12
    Public Const STYLE_ESSAY = "Q Essay"

    ' saves the current question type
    Dim questionType As String
    Dim xmlpath As String

    ' Prefix for the filename
    Const FILE_PREFIX = "Moodle_Questions_" 'this has an underscore obscured by the line


    ' Add a new paragraph with a specified style and text
    ' Inserted text is selected
    Public Sub AddParagraphOfStyle(aStyle, text)
        Dim myRange As Word.Range = Globals.ThisDocument.Application.Selection.Range
        With myRange
            .InsertParagraphBefore()
            '.Move Unit:=wdParagraph, Count:=1
            .Text = text '.InsertBefore(text)
            .Style = aStyle
            .Select()
        End With
    End Sub

    ' Moodle requires a dot (.) as a decimal separator. Thus, all comma separators need to
    ' be converted.
    Private Sub ConvertDecimalSeparator(ByVal aRange As Microsoft.Office.Interop.Word.Range)
        aRange.Find.Execute(FindText:=",", ReplaceWith:=".", _
        Format:=False, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
    End Sub

    ' Remove all formatting from the document
    Private Sub RemoveFormatting()
        Globals.ThisDocument.Application.Selection.WholeStory()
        Globals.ThisDocument.Application.Selection.Find.ClearFormatting()
    End Sub

    ' Count the number of paragraphs having the specified
    ' style in the defined range
    Function CountStylesInRange(aStyle As String, startPoint As Integer, endPoint As Integer) As Integer
        Dim aRange As Microsoft.Office.Interop.Word.Range
        Dim endP
        Dim counter
        aRange = getDocumentRange(startPoint, endPoint)
        endP = aRange.End  'store end point
        counter = 0

        With aRange.Find
            .ClearFormatting()
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Style = aStyle
            .Format = True
            Do While .Execute(Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop) = True
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
    Function StyleFound(aStyle As Microsoft.Office.Interop.Word.Style, aRange As Microsoft.Office.Interop.Word.Range) As Boolean
        With aRange.Find
            .ClearFormatting()
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Style = aStyle
            .Format = True
            .Execute(Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop)
        End With
        'MsgBox "Style: " & aStyle & " Found: " & aRange.Find.Found
        StyleFound = aRange.Find.Found
    End Function

    ' Removes answer weights from the selection
    Public Sub RemoveAnswerWeightsFromTheSelection()
        With Globals.ThisDocument.Application.Selection.Find
            .ClearFormatting()
            .Style = STYLE_ANSWERWEIGHT
            .text = ""
            .Replacement.text = ""
            .Forward = True
            .Format = True
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
        End With
    End Sub


    ' Checks the questionnaire.
    ' Returns true if everything is fine, otherwise false
    Function CheckQuestionnaire() As Boolean
        'return false if empty document
        If getCharacterCount() = 1 Then Return False

        Dim startOfQuestion, endOfQuestion, setEndPoint
        Dim isOK As Boolean

        isOK = True
        setEndPoint = False ' indicates whether the question end point should be set
        startOfQuestion = 0
        questionType = ""

        ' Check each paragraph at a time and specify needed tags
        For Each para As Paragraph In getParagraphs()

            ' Check if empty paragraph
            If para.Range.Text = vbCr Then
                para.Range.Delete() ' delete all empty paragraphs
                If questionType = "" Then questionType = para.Range.Style.NameLocal
            ElseIf para.Range.Style.NameLocal = STYLE_MULTICHOICEQ Or _
                   para.Range.Style.NameLocal = STYLE_MULTICHOICEQ_FIXANSWER Or _
                   para.Range.Style.NameLocal = STYLE_MATCHINGQ Or _
                   para.Range.Style.NameLocal = STYLE_MATCHINGQ_FIXANSWER Or _
                   para.Range.Style.NameLocal = STYLE_NUMERICALQ Or _
                   para.Range.Style.NameLocal = STYLE_SHORTANSWERQ Then

                If setEndPoint Then
                    endOfQuestion = para.Range.Start
                    isOK = CheckQuestion(startOfQuestion, endOfQuestion)
                    If isOK = False Then Exit For ' Exit if error is found
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
                    isOK = CheckQuestion(startOfQuestion, endOfQuestion)
                    If isOK = False Then Exit For ' Exit if error is found
                    startOfQuestion = para.Range.Start
                End If

                questionType = para.Range.Style.NameLocal

                isOK = CheckQuestion(startOfQuestion, para.Range.End)
                If isOK = False Then Exit For ' Exit if error is found
                startOfQuestion = para.Range.End
                questionType = "NOT_KNOWN"
                setEndPoint = False
            ElseIf para.Range.Style.NameLocal = STYLE_CORRECTANSWER And _
                   questionType = STYLE_NUMERICALQ Then
                ' Exit if error is found
                If CheckNumericAnswer(para.Range) = False Then Return False 'Exit Function
            End If

            ' Check if the end of document
            If para.Range.End = getRangeEnd() And _
            startOfQuestion <> para.Range.End Then
                isOK = CheckQuestion(startOfQuestion, getRangeEnd())
            End If

            If isOK = False Then Exit For ' Exit if error is found

        Next para

        'TODO not sure this makes sense, it will just skip the refresh
        If getCharacterCount() = 1 Then Return isOK


        moveCursorToEndOfDocument()
        Globals.ThisDocument.Application.ScreenRefresh()
        Return isOK
    End Function

    ' Checks whether the chosen question is valid
    ' Returns true if the question is OK, otherwise
    Function CheckQuestion(startPoint As Integer, endPoint As Integer) As Boolean
        Dim isOk As Boolean
        Dim rightCount, rightPairCount, leftPairCount, wordCount As Integer

        Dim aRange As Range

        aRange = getDocumentRange(startPoint, endPoint)
        aRange.Select()
        'MsgBox "See Range for specifying question type." & questionType & vbCr & _
        '      "Start: " & startPoint & " End: " & endPoint
        isOk = True 'no errors

        If questionType = STYLE_MULTICHOICEQ Or _
          questionType = STYLE_MULTICHOICEQ_FIXANSWER Then

            ' Check that there are right anwers specified
            rightCount = CountStylesInRange(STYLE_CORRECTANSWER, startPoint, endPoint)

            If rightCount = 0 Then
                aRange.Select()
                MsgBox("Error, no correct answer defined.", vbExclamation)
                isOk = False
            End If

        ElseIf questionType = STYLE_SHORTANSWERQ Then
            rightCount = CountStylesInRange(STYLE_SHORT_ANSWER, startPoint, endPoint)

            If rightCount = 0 Then
                aRange.Select()
                MsgBox("Error, no correct short answer is defined.", vbExclamation)
                isOk = False
            End If

        ElseIf questionType = STYLE_NUMERICALQ Then
            rightCount = CountStylesInRange(STYLE_SHORT_ANSWER, startPoint, endPoint)

            If rightCount = 0 Then
                aRange.Select()
                MsgBox("Error, no correct numerical answer is defined.", vbExclamation)
                isOk = False
            End If

            ' MATCHING QUESTION
        ElseIf questionType = STYLE_MATCHINGQ Or questionType = STYLE_MATCHINGQ_FIXANSWER Then

            ' Count the number of pairs
            rightPairCount = CountStylesInRange(STYLE_RIGHT_PAIR, startPoint, endPoint)
            leftPairCount = CountStylesInRange(STYLE_LEFT_PAIR, startPoint, endPoint)

            ' Too few pairs
            If leftPairCount < 3 Then
                aRange.Select()
                MsgBox("Error, there are not enough pairs for a matching question" & vbCr & _
                       "There must be at least 3 matching pairs. Please add more.", vbExclamation, "Error!")
                isOk = False
                ' Error -> the number of left and right pairs is different or zero
            ElseIf rightPairCount <> leftPairCount Then
                aRange.Select()
                MsgBox("Error, pairs are not correctly defined" & vbCr & _
                       "The number of left and right pairs is not equal.", vbExclamation, "Error!")
                isOk = False
            End If

        ElseIf questionType = STYLE_MISSINGWORDQ Then
            wordCount = CountStylesInRange(STYLE_BLANK_WORD, startPoint, endPoint)
            If wordCount <> 1 Then
                aRange.Select()
                MsgBox("There must be exactly one answer specified as a blank word." _
                + Chr(13) + Chr(13) + "To remove unnecessary markup, select a word(s) and press Ctrl+Space.", vbExclamation, "Error!")
                isOk = False
            End If

        ElseIf questionType = STYLE_TRUESTATEMENT Or _
               questionType = STYLE_FALSESTATEMENT Or _
               questionType = STYLE_ESSAY Then
            'nothing to check for answers to these ones. Figure out what the issue is with being last question in test and fix here?

            ' UNDEFINED QUESTION TYPE
        Else
            aRange.Select()
            MsgBox("Undefined Question type:" & questionType & vbCr & vbCr _
                   & "Illegal question is deleted.", vbExclamation, "Error!")
            aRange.Delete()
        End If

        'MsgBox questionType & "=" & isOk  'debugging output - show OK after each q checked.
        CheckQuestion = isOk
    End Function

    ' Check the numeric answer. Note, not checking all valid GIFT formats.
    Function CheckNumericAnswer(aRange As Range) As Boolean

        CheckNumericAnswer = True ' By default OK

        ' Search for the error margin separator
        aRange.Find.Execute(FindText:=":", Format:=False)

        If aRange.Find.Found = False And IsNumeric(aRange) = False Then
            Dim Response As Integer

            aRange.Select()
            Response = MsgBox("Is this a right numerical answer?" & vbCr & _
                       "Your answer: " & aRange.Characters.ToString, vbYesNo, "Correct?")  ' CPF not sure about converting the aRange to a string here...
            If Response = vbNo Then CheckNumericAnswer = False
        End If
    End Function

    ' Inserts text before the specified range. A new paragraph is inserted.
    Sub InsertTextBeforeRange(ByVal text As String, ByVal aRange As Range)
        With aRange
            .Style = STYLE_NORMAL
            .InsertBefore(text)
            .Move(Unit:=WdUnits.wdParagraph, Count:=1)
        End With

    End Sub

    ' Inserts text having trailing VbCr before the range
    Sub InsertQuestionEndTag(endPoint As Integer)
        Dim aRange As Range
        aRange = Globals.ThisDocument.Application.Range(endPoint - 1, endPoint)
        With aRange
            '            .InsertBefore(vbCr & TAG_QUESTION_END)
            .InsertBefore(vbCr)
            .Style = STYLE_NORMAL
            .Move(Unit:=WdUnits.wdParagraph, Count:=1)
        End With

    End Sub

    ' Inserts text at the end of the paragraph before the trailing VbCr
    Sub InsertAfterBeforeCR(ByVal text As String, ByVal aRange As Range)
        aRange.End = aRange.End - 1 ' insert text before cr
        With aRange
            .InsertAfter(text)
            .Style = STYLE_NORMAL
            .Move(Unit:=WdUnits.wdParagraph, Count:=1)
        End With
    End Sub

    ' Inserts text at the end of the paragraph
    Sub InsertAfterRange(ByVal text As String, aStyle As Style, ByVal aRange As Range)
        With aRange
            .EndOf(Unit:=WdUnits.wdParagraph, Extend:=WdMovementType.wdMove)
            .InsertParagraphBefore()
            .Move(Unit:=WdUnits.wdParagraph, Count:=-1)
            .Style = aStyle
            .InsertBefore(text)
            .Select()
        End With
    End Sub

    ' Set the answer weights of multiple choice questions.
    Public Sub SetAnswerWeights() ' aStyle, startPoint, endPoint)
        Dim startPoint, endPoint, rightScore, wrongScore, rightCount, wrongCount As Integer

        If getSelectionStyle() = STYLE_MULTICHOICEQ Or STYLE_MULTICHOICEQ_FIXANSWER Then
            startPoint = Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.Start
            rightCount = 0
            wrongCount = 0
            Globals.ThisDocument.Application.Selection.MoveDown(Unit:=WdUnits.wdParagraph, Count:=1)

            Do While getSelectionStyle() = STYLE_CORRECTANSWER Or _
                  getSelectionStyle() = STYLE_INCORRECTANSWER Or _
                  getSelectionStyle() = STYLE_FEEDBACK Or _
                  getSelectionStyle() = STYLE_ANSWERWEIGHT

                'Delete empty paragraphs
                If Globals.ThisDocument.Application.Selection.Paragraphs(1).Range = vbCr Then
                    Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.Delete() ' delete all empty paragraphs
                    ' Remove old answer weights
                ElseIf getSelectionStyle() = STYLE_ANSWERWEIGHT Then
                    With Globals.ThisDocument.Application.Selection.Find
                        .ClearFormatting()
                        .Style = STYLE_ANSWERWEIGHT
                        .Text = ""
                        .Replacement.Text = ""
                        .Forward = True
                        .Format = True
                        .Execute(Replace:=WdReplace.wdReplaceOne)
                    End With
                End If

                ' Count the number of right and wrong answers
                If getSelectionStyle() = STYLE_CORRECTANSWER Then
                    rightCount = rightCount + 1
                ElseIf getSelectionStyle() = STYLE_INCORRECTANSWER Then
                    wrongCount = wrongCount + 1
                End If

                If Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.End = Globals.ThisDocument.Application.Range.End Then
                    endPoint = Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.End
                    Exit Do
                Else
                    Globals.ThisDocument.Application.Selection.MoveDown(Unit:=WdUnits.wdParagraph, Count:=1)
                    endPoint = Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.Start
                End If
            Loop

            Dim QuestionRange As Range
            QuestionRange = Globals.ThisDocument.Application.Range(startPoint, endPoint)

            If rightCount < 1 Then
                QuestionRange.Select()
                MsgBox("No correct answer specified.", vbExclamation, "Error!")
            Else
                ' Calculate the right and wrong scores
                rightScore = Math.Round(100 / rightCount, 3)
                ' MODIFY the default scoring principle for wrong answers if necessary
                wrongScore = -rightScore

                AddAnswerWeights(QuestionRange, rightScore, wrongScore)
            End If
        Else
            MsgBox("Place the cursor on the question title" & vbCr & _
                   "of the Multiple Choice Question", vbExclamation, "Error!")
            ' Find the previous paragraph having the style of multiple choice question.
            With Globals.ThisDocument.Application.Selection.Find
                .ClearFormatting()
                .Text = ""
                .Style = STYLE_MULTICHOICEQ Or STYLE_MULTICHOICEQ_FIXANSWER
                .Forward = False
                .Format = True
                .MatchCase = False
                .Execute()
            End With
        End If
    End Sub

    ' Insert answer weights
    Private Sub AddAnswerWeights(ByVal aRange As Range, rightScore As Integer, wrongScore As Integer)
        ' Check each paragraph at a time and specify needed tags
        For Each para In aRange.Paragraphs
            ' Check if empty paragraph
            If para.Range = vbCr Then
                para.Range.Delete() ' delete all empty paragraphs
            ElseIf para.Range.Style = STYLE_CORRECTANSWER Then
                InsertAnswerWeight(rightScore, para.Range)
            ElseIf para.Range.Style = STYLE_INCORRECTANSWER Then
                InsertAnswerWeight(wrongScore, para.Range)
            End If
        Next para

    End Sub

    ' Inserts text at the end of the chapter before the trailing VbCr
    Private Sub InsertAnswerWeight(Score As Integer, ByVal aRange As Range)
        Dim startPoint As Integer
        Dim scoreString As String
        Dim newRange As Range
        startPoint = aRange.Start
        scoreString = "" & Score & "%"
        aRange.InsertBefore(scoreString)
        newRange = Globals.ThisDocument.Application.Range(Start:=startPoint, End:=startPoint + Len(scoreString))
        newRange.Style = STYLE_ANSWERWEIGHT
        'Moodle requires that the decimal separator is dot, not comma.
        ConvertDecimalSeparator(newRange)

    End Sub

    'LTG: This seems to be unused? left over from GIFT format days?
    Private Sub FindBlanks(aRange As Range)
        'Set aRange = Globals.ThisDocument.Application.Selection.Paragraphs(1).Range


        'ActiveDocument.Content
        Dim endPoint As Integer
        endPoint = aRange.End
        With aRange.Find
            .ClearFormatting()
            .Style = STYLE_BLANK_WORD
            Do While .Execute(FindText:="", Forward:=True, Format:=True) = True And _
                Globals.ThisDocument.Application.Selection.Range.End < endPoint And Globals.ThisDocument.Application.Range.End <> Globals.ThisDocument.Application.Selection.Range.End
                With .Parent
                    .InsertBefore("{=")
                    .InsertAfter("}")
                    '  .Move Unit:=wdParagraph, Count:=1
                    .Move(Unit:=WdUnits.wdWord, Count:=1)
                End With
            Loop
        End With
    End Sub



    Private Sub Convert2XML()


        ' Macro recorded on 21.12.2008 by Daniel Refresh Header (translation?)
        Globals.ThisDocument.Application.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
        Globals.ThisDocument.Application.Selection.Fields.Update()
        Globals.ThisDocument.Application.Selection.EndKey(Unit:=WdUnits.wdLine)
        Globals.ThisDocument.Application.Selection.MoveLeft(Unit:=WdUnits.wdCharacter, Count:=1)
        Globals.ThisDocument.Application.Selection.Fields.Update()
        Globals.ThisDocument.Application.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument


        'look for the folder containing .xml question patterns
        xmlpath = Globals.ThisDocument.Application.Path & "\xml-question\"
        If Not DirExists(xmlpath) Then
            If Globals.ThisDocument.Application.AttachedTemplate.Path <> "" Then
                xmlpath = Globals.ThisDocument.Application.AttachedTemplate.Path & "\xml-question\"
            Else
                xmlpath = Globals.ThisDocument.Application.Path & "\xml-question\"
            End If
            If Not DirExists(xmlpath) Then
                MsgBox("The xml-question\ folder is not found. Please keep this quiz question document in the original folder.", vbCritical, "Error")
                Exit Sub
            End If
        End If

        'choose the file name to save with
        Dim fd As Microsoft.Office.Core.FileDialog
        fd = Globals.ThisDocument.Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogSaveAs)
        '.FilterIndex = 2 fuer Word 2003, 14 fuer Word 2010
        fd.FilterIndex = 14
        fd.InitialFileName = FILE_PREFIX & Format(Now, "yyyymmdd") & ".xml"
        If fd.Show <> -1 Then Exit Sub

        Dim header As String
        header = Globals.ThisDocument.Application.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text


        '//*** save the file in utf-8 using stream ***//
        Dim objStream As ADODB.Stream
        'Create the stream
        objStream = CreateObject("ADODB.Stream")
        'Initialize the stream
        objStream.Open()
        'Reset the position and indicate the charactor encoding
        objStream.Position = 0
        objStream.Charset = "UTF-8"

        'specify XML version and that it is a quiz
        objStream.WriteText("<?xml version=""1.0""?><quiz>" & vbCr)
        'write the categories from the header
        objStream.WriteText("<question type=""category"">" & vbCr)
        objStream.WriteText("<category>" & vbCr)
        objStream.WriteText("<text>" & header & "</text>" & vbCr)
        objStream.WriteText("</category>" & vbCr)
        objStream.WriteText("</question>" & vbCr & vbCr)

        Dim dd As MSXML2.DOMDocument60
        Dim xmlnod As MSXML2.IXMLDOMNode
        'Dim xmlnodelist As MSXML2.IXMLDOMNodeList
        Dim para As Paragraph, paralookahead As Paragraph
        paralookahead = Nothing

        Dim rac, wac As Integer

        For Each para In Globals.ThisDocument.Application.Paragraphs '?handle each paragraph separately.
            dd = New MSXML2.DOMDocument60

            Select Case para.Style

                Case STYLE_SHORTANSWERQ
                    dd.load(xmlpath & "shortanswer.xml")
                    ProcessCommonTags(dd, para)
                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)
                    Do While (paralookahead.Style = STYLE_SHORT_ANSWER)
                        xmlnod.attributes.getNamedItem("fraction").text = "100"
                        xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        dd.documentElement.appendChild(xmlnod)

                        xmlnod = xmlnod.cloneNode(True)
                        paralookahead = paralookahead.Next
                        If paralookahead Is Nothing Then Exit Do
                    Loop

                Case STYLE_ESSAY
                    dd.load(xmlpath & "essay.xml")
                    ProcessCommonTags(dd, para)
                    'LTG: do I need to have Set paralookahead = para.Next here? why / why not?

                Case STYLE_NUMERICALQ
                    dd.load(xmlpath & "numerical.xml")
                    ProcessCommonTags(dd, para)
                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)

                    Do While (paralookahead.Style = STYLE_SHORT_ANSWER)
                        xmlnod.attributes.getNamedItem("fraction").text = "100"
                        xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        paralookahead = paralookahead.Next
                        If Not paralookahead Is Nothing Then
                            If (paralookahead.Style = STYLE_NUM_TOLERANCE) Then
                                xmlnod.selectSingleNode("tolerance").text = RemoveCR(paralookahead.Range.Text)
                                paralookahead = paralookahead.Next
                            Else
                                xmlnod.selectSingleNode("tolerance").text = "0"
                            End If
                        End If
                        dd.documentElement.appendChild(xmlnod)
                        xmlnod = xmlnod.cloneNode(True)
                        If paralookahead Is Nothing Then Exit Do
                    Loop

                Case STYLE_FALSESTATEMENT
                    dd.load(xmlpath & "false.xml")
                    ProcessCommonTags(dd, para)
                    paralookahead = para.Next

                Case STYLE_TRUESTATEMENT
                    dd.load(xmlpath & "true.xml")
                    ProcessCommonTags(dd, para)
                    paralookahead = para.Next

                Case STYLE_MULTICHOICEQ_FIXANSWER
                    dd.load(xmlpath & "multichoicefix.xml")
                    ProcessCommonTags(dd, para)

                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)

                    rac = 0
                    wac = 0
                    Do While (paralookahead.Style = STYLE_CORRECTANSWER) Or (paralookahead.Style = STYLE_INCORRECTANSWER)
                        If paralookahead.Style = STYLE_CORRECTANSWER Then
                            xmlnod.attributes.getNamedItem("fraction").text = "100"
                            rac = rac + 1
                        Else
                            xmlnod.attributes.getNamedItem("fraction").text = "0"
                            wac = wac + 1
                        End If
                        xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        dd.documentElement.appendChild(xmlnod)

                        xmlnod = xmlnod.cloneNode(True)
                        paralookahead = paralookahead.Next
                        If paralookahead Is Nothing Then Exit Do
                    Loop

                    If rac > 1 Then
                        ' multiple correct/incorrect answers
                        dd.documentElement.selectSingleNode("single").text = "false"
                        ' re-looping for setting multi-true-answer fractions
                        For Each mansw In dd.documentElement.selectNodes("answer")
                            With mansw.Attributes.getNamedItem("fraction")
                                If .text = 100 Then .text = Replace(100 / rac, ",", ".")
                                If .text = 0 Then .text = Replace(-100 / wac, ",", ".")
                            End With
                        Next mansw
                    End If

                Case STYLE_MULTICHOICEQ
                    dd.load(xmlpath & "multichoicevar.xml")
                    ProcessCommonTags(dd, para)

                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)
                    rac = 0 'right answer choices
                    wac = 0 'wrong answer choices
                    ' loop breaks wrongly because of Feedback Style
                    Do While (paralookahead.Style = STYLE_CORRECTANSWER) Or (paralookahead.Style = STYLE_INCORRECTANSWER)
                        If paralookahead.Style = STYLE_CORRECTANSWER Then
                            xmlnod.attributes.getNamedItem("fraction").text = "100"
                            rac = rac + 1
                        Else
                            xmlnod.attributes.getNamedItem("fraction").text = "0"
                            wac = wac + 1
                        End If
                        xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)

                        paralookahead = paralookahead.Next
                        ' Feedback Style processing here
                        If paralookahead.Style = STYLE_FEEDBACK Then
                            ' Set XML <feedback> text
                            xmlnod.selectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.Text)
                            paralookahead = paralookahead.Next
                        End If

                        dd.documentElement.appendChild(xmlnod)
                        xmlnod = xmlnod.cloneNode(True)

                        If paralookahead Is Nothing Then Exit Do
                    Loop

                    If rac > 1 Then
                        ' multiple correct/incorrect answers
                        dd.documentElement.selectSingleNode("single").text = "false"
                        ' re-looping for setting multi-true-answer fractions
                        For Each mansw In dd.documentElement.selectNodes("answer")
                            With mansw.Attributes.getNamedItem("fraction")
                                If .text = 100 Then .text = Replace(100 / rac, ",", ".") 'original
                                If .text = 0 Then .text = Replace(-100 / wac, ",", ".") 'original
                                'If .text = 100 Then .text = Round(100 / rac, 5) 'sd 2010 für Moodle 2.0
                                'If .text = 0 Then .text = Round(-100 / wac, 5) 'sd 2010 für moodle 2.0
                            End With
                        Next mansw
                    End If

                Case STYLE_MATCHINGQ
                    dd.load(xmlpath & "matchingvar.xml")
                    ProcessCommonTags(dd, para)

                    ' processing each <subquestion>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("subquestion")
                    dd.documentElement.removeChild(xmlnod)
                    Do While (paralookahead.Style = STYLE_LEFT_PAIR) Or (paralookahead.Style = STYLE_RIGHT_PAIR)
                        If paralookahead.Style = STYLE_LEFT_PAIR Then
                            xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        Else
                            xmlnod.selectSingleNode("answer").selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)

                            dd.documentElement.appendChild(xmlnod)
                            xmlnod = xmlnod.cloneNode(True)
                        End If
                        paralookahead = paralookahead.Next
                        If paralookahead Is Nothing Then Exit Do
                    Loop

                Case STYLE_MATCHINGQ_FIXANSWER
                    dd.load(xmlpath & "matchingfix.xml")
                    ProcessCommonTags(dd, para)

                    ' processing each <subquestion>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("subquestion")
                    dd.documentElement.removeChild(xmlnod)
                    Do While (paralookahead.Style = STYLE_LEFT_PAIR) Or (paralookahead.Style = STYLE_RIGHT_PAIR)
                        If paralookahead.Style = STYLE_LEFT_PAIR Then
                            xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        Else
                            xmlnod.selectSingleNode("answer").selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)

                            dd.documentElement.appendChild(xmlnod)
                            xmlnod = xmlnod.cloneNode(True)
                        End If
                        paralookahead = paralookahead.Next
                        If paralookahead Is Nothing Then Exit Do
                    Loop

                Case STYLE_MISSINGWORDQ
                    dd.load(xmlpath & "shortanswer.xml")
                    Dim theChar As Range
                    Dim misword As String
                    misword = ""
                    For Each theChar In para.Range.Characters
                        If theChar.Style = STYLE_BLANK_WORD Then misword = misword & theChar.Text
                    Next theChar
                    ProcessCommonTags(dd, para)
                    dd.documentElement.selectSingleNode("name").selectSingleNode("text").text = Replace( _
                      dd.documentElement.selectSingleNode("name").selectSingleNode("text").text, misword, "__________")

                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    xmlnod.attributes.getNamedItem("fraction").text = "100"
                    xmlnod.selectSingleNode("text").text = misword

                Case STYLE_COMMENT
                    Dim Comment As String
                    Comment = "<!-- " & RemoveCR(para.Range.Text) & " -->"
                    objStream.WriteText(Comment & vbCr & vbCr)
                    dd.loadXML("")

                Case Else
                    dd.loadXML("")
            End Select


            If Not paralookahead Is Nothing Then
                If (paralookahead.Style = STYLE_QUESTIONNAME) Then
                    xmlnod = dd.documentElement.selectSingleNode("name")
                    dd.documentElement.removeChild(xmlnod)
                    xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                    dd.documentElement.appendChild(xmlnod)
                    xmlnod = xmlnod.cloneNode(True)
                    paralookahead = paralookahead.Next
                End If

            End If
            If Not paralookahead Is Nothing Then '**seems to be setting generalfeedback for any feedback tag...
                ' CPF commented out
                '          If (paralookahead.Style = STYLE_FEEDBACK) Then
                '             Set xmlnod = dd.documentElement.SelectSingleNode("generalfeedback")
                '             dd.documentElement.RemoveChild xmlnod
                '             xmlnod.SelectSingleNode("text").text = RemoveCR(paralookahead.Range.text)
                '             dd.documentElement.appendChild xmlnod
                '             Set xmlnod = xmlnod.CloneNode(True)
                '             Set paralookahead = paralookahead.Next
                '          End If
            End If

            If dd.xml <> "" Then objStream.WriteText(dd.xml & vbCr)

            dd = Nothing
        Next para

        objStream.WriteText("</quiz>")
        'Save the stream to a file
        objStream.SaveToFile(FileName:=fd.SelectedItems(1), Options:=ADODB.SaveOptionsEnum.adSaveCreateOverWrite)

    End Sub

    'This is called as the first processing task for each question. It
    Private Sub ProcessCommonTags(dd As MSXML2.DOMDocument60, para As Paragraph)
        ' processing <name> '
        dd.documentElement.selectSingleNode("name") _
        .selectSingleNode("text").text = RemoveCR(para.Range.Text)

        ' processing <questiontext> '
        dd.documentElement.selectSingleNode("questiontext") _
        .selectSingleNode("text").text = XSLT_Range(para.Range, "FormattedText.xslt")

        '
        If Not XSLT_Range(para.Range, "PictureName.xslt") = "" Then 'if it is NOT null/empty

            Dim header As String
            Dim stringlength As Long
            header = Globals.ThisDocument.Application.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text
            stringlength = Len(header)
            header = Left(header, stringlength - 1)

            'processing <image>'
            dd.documentElement.selectSingleNode("image").text = "Images_forQuizQuestions/" & header & Right(XSLT_Range(para.Range, "PictureName.xslt"), 4)
            'dd.documentElement.SelectSingleNode("image").text = Mid(XSLT_Range(para.Range, "PictureName.xslt"), 10) (commented out: Rohrer)'
            'processing <image_base64>'
            dd.documentElement.selectSingleNode("image_base64").text = XSLT_Range(para.Range, "Picture.xslt")
        End If




    End Sub


    Private Function XSLT_Range(textrange As Range, xsltfilename As String) As String
        Dim xsldoc As New MSXML2.FreeThreadedDOMDocument60
        xsldoc.load(xmlpath & xsltfilename)
        Dim xslt As New MSXML2.XSLTemplate60
        xslt.stylesheet = xsldoc
        Dim xsltProcessor As MSXML2.IXSLProcessor
        xsltProcessor = xslt.createProcessor
        Dim d As New MSXML2.DOMDocument60
        Dim s As String
        d.loadXML(textrange.XML) '!!! Bug in Word 2010 when file is created from a template (Textrange.xml can not be read)
        xsltProcessor.input = d
        xsltProcessor.transform()
        s = xsltProcessor.output

        xsltProcessor = Nothing
        xslt = Nothing
        xsldoc = Nothing
        XSLT_Range = s
    End Function


    Private Function RemoveCR(str As String) As String
        str = Replace(str, vbCr, "")
        RemoveCR = Trim(Globals.ThisDocument.Application.CleanString(str))
    End Function


    Function DirExists(ByVal sDirName As String) As Boolean
        On Error Resume Next
        DirExists = (GetAttr(sDirName) And vbDirectory) = vbDirectory
        Err.Clear()
    End Function

    Private Function getSelectionStyle() As String
        Return CType(Globals.ThisDocument.Application.Selection.Paragraphs.Style, Word.Style).NameLocal
    End Function

    Private Sub setSelectionStyle(theStyle As String)
        Globals.ThisDocument.Application.Selection.Paragraphs.Style = theStyle
    End Sub
    Private Function getSelectionRange() As Range
        Return Globals.ThisDocument.Application.Selection.Range
    End Function

    Private Function getCharacterCount() As Integer
        Return Globals.ThisDocument.Application.ActiveDocument.Characters.Count
    End Function

    Private Function getParagraphs() As Paragraphs
        Return Globals.ThisDocument.Application.ActiveDocument.Paragraphs
    End Function

    Private Function getRangeEnd() As Integer
        Return Globals.ThisDocument.Application.ActiveDocument.Range.End
    End Function

    Private Function getDocumentRange(startPoint As Integer, endPoint As Integer) As Microsoft.Office.Interop.Word.Range
        Return Globals.ThisDocument.Application.ActiveDocument.Range(startPoint, endPoint)
    End Function

    Private Sub moveCursorToEndOfDocument()
        Globals.ThisDocument.Application.Selection.EndKey(WdUnits.wdStory, Nothing)
    End Sub

    Private Sub moveCursorToStartOfDocument()
        Globals.ThisDocument.Application.Selection.HomeKey(WdUnits.wdStory, Nothing)
    End Sub

End Class
