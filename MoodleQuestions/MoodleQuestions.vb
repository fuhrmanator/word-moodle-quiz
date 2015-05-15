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
'FIXED Style "A Short Answer" is missing from .DOCM
'FIXED find a deployment site for Project Publishing. Google Drive won't work because it doesn't have clean URLs for directories.
'TODO Try using style content to indicate [shuffled] questions (rather than arbitrary colors). Numbering allows inserting text after the 1. (e.g., 1. [S] for shuffled)
'TODO For which questions is answer feedback valid? Do we need different feedback word styles?
'TODO Figure out what "Question Name" button is supposed to do
'TODO Add "Question Feedback" button (different from Answer feedback)
'TODO Understand Numerical Questions: "Q Numerical" is followed by "Short Answer" in the v21 template. Should we make a "A Numerical" for consistency? Might impact "Check Layout" function.
'TODO Fix XML Export for all question types
'TODO Selecting a missing word and using the button also selects the space after the word, which cause a problem in <questiontext/text>
'TODO The carriage return doesn't work for context change


Imports Microsoft.Office.Interop.Word
Imports stdole
Imports MSXML2
Imports System.Runtime.InteropServices
Imports System.Collections
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView

<Runtime.InteropServices.ComVisible(True)> _
Public Class MoodleQuestions
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI
    Dim enabled As Boolean


    Public Sub New()
        enabled = False
    End Sub


    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("MoodleQuestions.MoodleQuestions.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        Globals.ThisDocument.ribbon = Me.ribbon
        Me.ribbon.ActivateTab("MoodleQuestions") 'Make Moodle Questions toolbar active on startup
        updateVersionInfo()
        With Globals.ThisDocument.Application.Selection.Range
            .Move(Unit:=WdUnits.wdParagraph, Count:=+1)
            .Select()
        End With
        
        CheckStyle()
    End Sub


    Public Function OnLoadImage(imageId As String) As IPictureDisp
        Dim tempImage As stdole.IPictureDisp = Nothing
        'load image from resources file
        tempImage = Microsoft.VisualBasic.Compatibility.VB6.Support.ImageToIPicture(My.Resources.RibbonIcons.ResourceManager.GetObject(imageId))
        Return tempImage

    End Function

    Public Sub CheckStyle()
        Dim styleCollection As New Microsoft.VisualBasic.Collection()
        Dim defaultstyleCollection As New Microsoft.VisualBasic.Collection()

        styleCollection.Add("Q Multi Choice")
        styleCollection.Add("Q Multi Choice FixAnswer")
        styleCollection.Add("Q Matching")
        styleCollection.Add("A Correct Choice")
        styleCollection.Add("A Incorrect Choice")
        styleCollection.Add("Q True Statement")
        styleCollection.Add("Q False Statement")
        styleCollection.Add("Q Missing Word")
        styleCollection.Add("A Feedback")
        styleCollection.Add("Q Category")
        styleCollection.Add("Q Short Answer")
        styleCollection.Add("A Short Answer")
        styleCollection.Add("Q Numerical")
        styleCollection.Add("Q Matching FixAnswer")
        styleCollection.Add("A Matching Left")
        styleCollection.Add("A Matching Right")
        styleCollection.Add("MissingWord")
        styleCollection.Add("Questionname")
        styleCollection.Add("Q Essay")
        styleCollection.Add("A Feedback FS")
        styleCollection.Add("A Feedback TS")

        For Each styleName In Globals.ThisDocument.Styles
            defaultstyleCollection.Add(styleName.NameLocal)
        Next

        Dim found As Boolean = False
        For Each item As String In styleCollection
            For Each itemDef As String In defaultstyleCollection
                If item = itemDef Then
                    found = True
                    Exit For
                End If
            Next
            If found = False Then
                MsgBox("Missing style: " + item)
                ' Exit For
            End If
            found = False
        Next

    End Sub

    '''''BUTTON callbacks
    ' Add Multiple Choice Question to the end of the active document
    Public Sub displayVersionInfo(ByVal control As Office.IRibbonControl)
        MsgBox("About MoodleQuestions..." & vbCrLf & "Published version: " & VERSION_INFO & vbCrLf & SOURCE_CODE_URL)
    End Sub


    ' Change the button states in the Ribbon
    Public Function GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
        System.Diagnostics.Debug.WriteLine("caught GetEnabled for " + control.Id)
        Dim isEnabled As Boolean = False
        Dim selectionStyleName As String = getSelectionStyleName()
        Select Case control.Id
            Case "shuffleanswers"
                isEnabled = (selectionStyleName = STYLE_MULTICHOICEQ Or _
                             selectionStyleName = STYLE_MULTICHOICEQ_FIXANSWER Or _
                             selectionStyleName = STYLE_MATCHINGQ)
            Case "MarkTrueFalse"
                isEnabled = (selectionStyleName = STYLE_CORRECT_MC_ANSWER Or _
                             selectionStyleName = STYLE_INCORRECT_MC_ANSWER Or _
                             selectionStyleName = STYLE_TRUESTATEMENT Or _
                             selectionStyleName = STYLE_FALSESTATEMENT)
            Case "MarkMissingWord"
                isEnabled = (selectionStyleName = STYLE_MISSINGWORDQ)
            Case "pasteImage"
                isEnabled = (selectionStyleName = STYLE_MULTICHOICEQ Or _
                             selectionStyleName = STYLE_CORRECT_MC_ANSWER Or _
                            selectionStyleName = STYLE_INCORRECT_MC_ANSWER)
            Case "questionTitle"
                isEnabled = (getSelectionStyleName() = STYLE_MULTICHOICEQ Or _
                             getSelectionStyleName() = STYLE_FEEDBACK Or _
                             isSelectionNormalStyle())
            Case "feedback"
                isEnabled = (isSelectionNormalStyle())
        End Select
        Return isEnabled
    End Function


    ' Add Multiple Choice Question to the end of the active document
    Public Sub AddMultipleChoiceQText()
        InsertParagraphAfterCurrentParagraph("Insert Multiple Choice Question here", "Q Multi Choice")
        InsertParagraphAfterCurrentParagraph("Insert correct answer here", "A Correct Choice")
        InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a correct answer here", "A Feedback")
        InsertParagraphAfterCurrentParagraph("Insert incorrect answer here", "A Incorrect Choice")
        InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is an incorrect answer here", "A Feedback")
        InsertParagraphAfterCurrentParagraph("Insert incorrect answer here", "A Incorrect Choice")
        AddParagraphOfStyle(STYLE_FEEDBACK, "Insert feedback explaining why this is an incorrect answer here")

    End Sub


    Public Sub AddMultipleChoiceQ(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            AddMultipleChoiceQText()
        Else

            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                AddMultipleChoiceQText()
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_MULTICHOICEQ, "Insert Multiple Choice Question here", max)
                InsertParagraphAfterCurrentParagraph("Insert correct answer here", "A Correct Choice")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a correct answer here", "A Feedback")
                InsertParagraphAfterCurrentParagraph("Insert incorrect answer here", "A Incorrect Choice")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is an incorrect answer here", "A Feedback")
                InsertParagraphAfterCurrentParagraph("Insert incorrect answer here", "A Incorrect Choice")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is an incorrect answer here", "A Feedback")
            Else 'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_MULTICHOICEQ, "Insert Multiple Choice Question here", min)
                InsertParagraphAfterCurrentParagraph("Insert correct answer here", "A Correct Choice")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a correct answer here", "A Feedback")
                InsertParagraphAfterCurrentParagraph("Insert incorrect answer here", "A Incorrect Choice")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is an incorrect answer here", "A Feedback")
                InsertParagraphAfterCurrentParagraph("Insert incorrect answer here", "A Incorrect Choice")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is an incorrect answer here", "A Feedback")
            End If
        End If
    End Sub

    Public Sub AddCategoryQ(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("Question_Category/Question_Subcategory", "Q Category")
            '  AddParagraphOfStyle(STYLE_CATEGORYQ, "Question_Category/Question_Subcategory")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("Question_Category/Question_Subcategory", "Q Category")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_CATEGORYQ, "Question_Category/Question_Subcategory", max)
            Else 'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_CATEGORYQ, "Question_Category/Question_Subcategory", min)
            End If
        End If
    End Sub

    ' Add Matching Question to the end of the active document
    Public Sub AddMatchingQ(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("Insert Matching Question", "Q Matching")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("Insert Matching Question", "Q Matching")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_MATCHINGQ, "Insert Matching Question", max)
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_MATCHINGQ, "Insert Matching Question", min)
            End If
        End If
    End Sub

    ' Add Numerical Question to the end of the active document
    Public Sub AddNumericalQ(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("Insert Numerical Question", "Q Numerical")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("Insert Numerical Question", "Q Numerical")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_NUMERICALQ, "Insert Numerical Question", max)
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_NUMERICALQ, "Insert Numerical Question", min)
            End If
        End If
    End Sub


    ' Add Short Answer Question to the end of the active document
    Public Sub AddShortAnswerQ(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("Insert Short Answer Question", "Q Short Answer")
            InsertParagraphAfterCurrentParagraph("Insert Short Answer here", "A Short Answer")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("Insert Short Answer Question", "Q Short Answer")
                InsertParagraphAfterCurrentParagraph("Insert Short Answer here", "A Short Answer")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_SHORTANSWERQ, "Insert Short Answer Question", max)
                InsertParagraphAfterCurrentParagraph("Insert Short Answer here", "A Short Answer")
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_SHORTANSWERQ, "Insert Short Answer Question", min)
                InsertParagraphAfterCurrentParagraph("Insert Short Answer here", "A Short Answer")
            End If
        End If
    End Sub

    ' Add Missing Word Question
    Public Sub AddMissingWordQ(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("Insert Missing Word Question. Then select the missing word!", "Q Missing Word")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("Insert Missing Word Question. Then select the missing word!", "Q Missing Word")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_MISSINGWORDQ, "Insert Missing Word Question. Then select the missing word!", max)
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_MISSINGWORDQ, "Insert Missing Word Question. Then select the missing word!", min)
            End If
        End If
    End Sub

    ' Add an Essay
    Public Sub AddEssay(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("Insert An Essay question here (an Open Question). [This can not be the last question in the document.]", "Q Essay")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("Insert An Essay question here (an Open Question). [This can not be the last question in the document.]", "Q Essay")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_ESSAY, "Insert An Essay question here (an Open Question). [This can not be the last question in the document.]", max)
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_ESSAY, "Insert An Essay question here (an Open Question). [This can not be the last question in the document.]", min)
            End If
        End If
    End Sub
    Public Sub ToggleMissingWord(ByVal control As Office.IRibbonControl)
        ' Only applies to questions of STYLE_MISSINGWORDQ
        If (getSelectionStyleName() = STYLE_MISSINGWORDQ) Then
            ' get only the first word of the selection
            Dim aRange As Microsoft.Office.Interop.Word.Range = getDocumentSelectionRange()
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

    'TODO make sure only answers that can get feedback are allowed
    Public Sub AddAnswerFeedback(ByVal control As Office.IRibbonControl)
        If getSelectionStyleName() = STYLE_CORRECT_MC_ANSWER Or _
           getSelectionStyleName() = STYLE_INCORRECT_MC_ANSWER Or _
           getSelectionStyleName() = STYLE_TRUESTATEMENT Or _
           getSelectionStyleName() = STYLE_FALSESTATEMENT Or _
           getSelectionStyleName() = STYLE_SHORT_ANSWER Then
            InsertParagraphAfterCurrentParagraph("Insert feedback of the previous choice or answer here.", _
                             STYLE_FEEDBACK)
            MsgBox("Feedback is")

        Else 'Error: Give Instructions:
            MsgBox("Feedback is placed at the end of the last possible response. " & vbCr & _
                   "It doesn't work for True/False questions." & vbCr & _
                   "Place the cursor on top of the question or answer you are giving feedback for.", vbExclamation)
        End If
    End Sub
    ' Add tolerance
    Public Sub AddNumericalTolerance(ByVal control As Office.IRibbonControl)
        If getSelectionStyleName() = STYLE_SHORT_ANSWER Then
            InsertParagraphAfterCurrentParagraph("Replace me with Tolerance for the answer as a Decimal. Eg: 0.01", _
                             STYLE_NUM_TOLERANCE)
        Else 'Error: Give Instructions:
            MsgBox(" " & vbCr & _
                   "Place the cursor at the end of the numerical answer.", vbExclamation)
        End If
    End Sub

    ' Add QuestionName / Question Title
    Public Sub AddQuestionTitle(ByVal control As Office.IRibbonControl)
        If getSelectionStyleName() = STYLE_ANSWERWEIGHT Or _
           getSelectionStyleName() = STYLE_SHORTANSWERQ Or _
           getSelectionStyleName() = STYLE_MISSINGWORDQ Or _
           getSelectionStyleName() = STYLE_CORRECT_MC_ANSWER Or _
           getSelectionStyleName() = STYLE_NUM_TOLERANCE Or _
           getSelectionStyleName() = STYLE_INCORRECT_MC_ANSWER Or _
           getSelectionStyleName() = STYLE_TRUESTATEMENT Or _
           getSelectionStyleName() = STYLE_SHORT_ANSWER Or _
           getSelectionStyleName() = STYLE_FALSESTATEMENT Or _
           getSelectionStyleName() = STYLE_RIGHT_MATCH Or _
           getSelectionStyleName() = STYLE_BLANK_WORD Then
            InsertParagraphAfterCurrentParagraph("Add a question title.", STYLE_QUESTIONNAME)
        Else 'Error: Give Instructions:
            MsgBox("Feedback to insert at the end of the last response selected. " & vbCr & _
                   "The title must appear before the feedback" & vbCr & _
                   "Place the cursor at the end of the last line selected", vbExclamation)
        End If
    End Sub


    ' Add a true statement of the true-false question
    Public Sub AddTrueStatement(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("True-false question: insert a TRUE statement here (not at the end of the document)", "Q True Statement")
            InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a True statement here", "A Feedback TS")
            InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a not False statement here", "A Feedback FS")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("True-false question: insert a TRUE statement here (not at the end of the document)", "Q True Statement")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a True statement here", "A Feedback TS")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a False statement here", "A Feedback FS")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_TRUESTATEMENT, "True-false question: insert a TRUE statement here (not at the end of the document)", max)
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a True statement here", "A Feedback TS")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a False statement here", "A Feedback FS")
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_TRUESTATEMENT, "True-false question: insert a TRUE statement here (not at the end of the document)", min)
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a True statement here", "A Feedback TS")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a False statement here", "A Feedback FS")
            End If
        End If
    End Sub

    ' Add a false statement of the true-false question
    Public Sub AddFalseStatement(ByVal control As Office.IRibbonControl)
        If isSelectionNormalStyle() Then
            InsertParagraphAfterCurrentParagraph("True-false question: insert a FALSE statement here (not at the end of the document)", "Q False Statement")
            InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a True statement here", "A Feedback TS")
            InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a False statement here", "A Feedback FS")
        Else
            Dim min As Integer = Math.Min(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))
            Dim max As Integer = Math.Max(StyleFound(STYLE_CATEGORYQ), QuestionStyleFound(styleList))

            'both styles not found
            If StyleFound(STYLE_CATEGORYQ) = -1 And QuestionStyleFound(styleList) = -1 Then
                moveCursorToEndOfDocument()
                InsertParagraphAfterCurrentParagraph("True-false question: insert a FALSE statement here (not at the end of the document)", "Q False Statement")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a True statement here", "A Feedback TS")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a False statement here", "A Feedback FS")
                'one of the two style found
            ElseIf min = -1 Then
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_FALSESTATEMENT, "True-false question: insert a FALSE statement here (not at the end of the document)", max)
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a True statement here", "A Feedback TS")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a False statement here", "A Feedback FS")
            Else  'both styles found
                InsertParagraphOfStyleInSelectedRangeBefore(STYLE_FALSESTATEMENT, "True-false question: insert a FALSE statement here (not at the end of the document)", min)
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is not a True statement here", "A Feedback TS")
                InsertParagraphAfterCurrentParagraph("Insert feedback explaining why this is a False statement here", "A Feedback FS")
            End If
        End If

    End Sub

    ' Add a comment
    'Public Sub AddComment(ByVal control As Office.IRibbonControl)
    '    AddParagraphOfStyle(STYLE_COMMENT, "")
    'End Sub
    Public Sub PasteImage(ByVal control As Office.IRibbonControl)
        '  Adds an image from the clipboard into a question.
        '  Globals.ThisDocument.Application.Selection.Paragraphs.Style = TYLE_SHORTANSWERQ
        '  TODO for test: tester tous les types d'extension d'images

        With Globals.ThisDocument.Application.Selection
            If getSelectionStyleName() = STYLE_SHORTANSWERQ Or _
               getSelectionStyleName() = STYLE_MISSINGWORDQ Or _
               getSelectionStyleName() = STYLE_MULTICHOICEQ Or _
               getSelectionStyleName() = STYLE_MATCHINGQ Or _
               getSelectionStyleName() = STYLE_NUMERICALQ Or _
               getSelectionStyleName() = STYLE_TRUESTATEMENT Or _
               getSelectionStyleName() = STYLE_FALSESTATEMENT Or _
               getSelectionStyleName() = STYLE_MULTICHOICEQ_FIXANSWER Or _
               getSelectionStyleName() = STYLE_MATCHINGQ_FIXANSWER Or _
               getSelectionStyleName() = STYLE_CORRECT_MC_ANSWER Or _
               getSelectionStyleName() = STYLE_INCORRECT_MC_ANSWER Then
                Globals.ThisDocument.Application.Options.ReplaceSelection = False
                .TypeText(Text:=(" " & Chr(11)))
                If (Not Clipboard.ContainsText) Then
                    .Paste()  ' don't paste if clipboard is empty, else exception
                Else
                    MsgBox("Clipboard does not contain an image. ", vbExclamation)
                End If
            Else 'Error - give instructions:
                MsgBox("Pastes an image from the Clipboard. " & vbCr & _
                       "Place the cursor at the end of the question. ", vbExclamation)
            End If
        End With
    End Sub

    Public Sub ToggleAnswer(ByVal control As Office.IRibbonControl)
        'Toggles MCQ answer (right-wrong) or switches true and false statements.
        Dim theStyle As String = getSelectionStyleName()

        If theStyle = STYLE_CORRECT_MC_ANSWER Then
            setSelectionParagraphStyle(STYLE_INCORRECT_MC_ANSWER)
        ElseIf theStyle = STYLE_INCORRECT_MC_ANSWER Then
            setSelectionParagraphStyle(STYLE_CORRECT_MC_ANSWER)
        ElseIf theStyle = STYLE_TRUESTATEMENT Then
            setSelectionParagraphStyle(STYLE_FALSESTATEMENT)
        ElseIf theStyle = STYLE_FALSESTATEMENT Then
            setSelectionParagraphStyle(STYLE_TRUESTATEMENT)
        Else 'Error: give instructions:
            MsgBox("This command toggles a statement from True to False." & vbCr & _
                   "Cursor must be on an answer for Multiple Choice" & vbCr & _
                   "or on a True or False statement.", vbExclamation)
        End If
    End Sub

    Public Sub ChangeShuffleanswerTrueFalse(ByVal control As Office.IRibbonControl)

        If getSelectionStyleName() = STYLE_MATCHINGQ Then
            setSelectionParagraphStyle(STYLE_MATCHINGQ_FIXANSWER)
        ElseIf getSelectionStyleName() = STYLE_MATCHINGQ_FIXANSWER Then
            setSelectionParagraphStyle(STYLE_MATCHINGQ)
        ElseIf getSelectionStyleName() = STYLE_MULTICHOICEQ Then
            setSelectionParagraphStyle(STYLE_MULTICHOICEQ_FIXANSWER)
        ElseIf getSelectionStyleName() = STYLE_MULTICHOICEQ_FIXANSWER Then
            setSelectionParagraphStyle(STYLE_MULTICHOICEQ)

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

    Public VERSION_INFO As String = "unknown"
    Public Const SOURCE_CODE_URL As String = "https://code.google.com/p/word-moodle-quiz/"
    ' General purpose styles.
    Public Const STYLE_NORMAL = Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleNormal

    Public Const STYLE_FEEDBACK = "A Feedback"
    Public Const STYLE_ANSWERWEIGHT = "A Weight"

    Public Const STYLE_CATEGORYQ = "Q Category"
    Public Const STYLE_SHORTANSWERQ = "Q Short Answer"
    Public Const STYLE_MULTICHOICEQ = "Q Multi Choice"
    Public Const STYLE_MATCHINGQ = "Q Matching"
    Public Const STYLE_NUMERICALQ = "Q Numerical"
    Public Const STYLE_MISSINGWORDQ = "Q Missing Word"
    Public Const STYLE_TRUESTATEMENT = "Q True Statement"
    Public Const STYLE_FALSESTATEMENT = "Q False Statement"
    Public Const STYLE_CORRECT_MC_ANSWER = "A Correct Choice"
    Public Const STYLE_INCORRECT_MC_ANSWER = "A Incorrect Choice"
    Public Const STYLE_SHORT_ANSWER = "A Short Answer"
    Public Const STYLE_LEFT_MATCH = "A Matching Left"
    Public Const STYLE_RIGHT_MATCH = "A Matching Right"
    Public Const STYLE_BLANK_WORD = "MissingWord"  'this string used in an XSLT template
    ' Public Const STYLE_COMMENT = "Comment"
    ' Supplement(ed by) Daniel
    Public Const STYLE_MULTICHOICEQ_FIXANSWER = "Q Multi Choice FixAnswer"
    Public Const STYLE_MATCHINGQ_FIXANSWER = "Q Matching FixAnswer"
    Public Const STYLE_NUM_TOLERANCE = "Num Tolerance"
    Public Const STYLE_QUESTIONNAME = "Questionname"
    'from v12
    Public Const STYLE_ESSAY = "Q Essay"
    '#modif feedback
    Public Const STYLE_FEEDBACK_FS = "A Feedback FS"
    Public Const STYLE_FEEDBACK_TS = "A Feedback TS"

    ' saves the current question type
    Dim questionType As String
    Dim xmlpath As String

    ' Prefix for the filename
    Const FILE_PREFIX = "Moodle_Questions_" 'this has an underscore obscured by the line

    Dim styleList As String() = {STYLE_MULTICHOICEQ, STYLE_MATCHINGQ, STYLE_SHORTANSWERQ, STYLE_ESSAY, STYLE_TRUESTATEMENT,
     STYLE_FALSESTATEMENT, STYLE_MISSINGWORDQ, STYLE_NUMERICALQ}

    ' Add a new paragraph with a specified style and text
    ' Inserted text is selected
    Public Sub AddParagraphOfStyle(aStyle, text)
        Dim myRange As Word.Range = Globals.ThisDocument.Application.Selection.Range
        Try
            With myRange
                .InsertParagraphAfter() '.InsertBefore(text)
                .Move(Unit:=WdUnits.wdParagraph, Count:=1)
                .Style = aStyle
                .Text = text
                .Select()
            End With
        Catch ex As System.Runtime.InteropServices.COMException
            MessageBox.Show(aStyle + " style doesn't exist") 'ex.Message)
        End Try
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
    ' Check if the Category style is found in the range 
    Function StyleFound(aStyle) As Integer
        Dim rng As Word.Range
        rng = Globals.ThisDocument.Application.Selection.Range
        rng.Start = rng.Start + 10

        With rng.Find
            .ClearFormatting()
            .Style = aStyle
            .Forward = True
            .Format = True
            .Execute()
        End With
        If rng.Find.Found Then
            Return rng.End
        Else
            Return -1
        End If
    End Function


    'Check if the specified style is found in the range
    Function QuestionStyleFound(ByVal styleList As String()) As Integer
        Dim rng As Word.Range

        Dim find As Integer = 0
        Dim rangeFound(7) As Integer
        Dim min As Integer = 3000
        Dim i As Integer = 0
        For Each element As String In styleList
            rng = Globals.ThisDocument.Application.Selection.Range
            rng.Start = rng.Start + 1
            With rng.Find
                .ClearFormatting()
                .Style = element
                .Forward = True
                .Format = True
                .Execute()
            End With
            If rng.Find.Found Then
                find = rng.End
            Else
                find = -1
            End If
            rangeFound(i) = find
            i += 1
        Next

        For Each element As Integer In rangeFound
            If element > 0 Then
                find = Math.Min(min, element)
                min = find
            End If
        Next
        Return find
    End Function

    ' Removes answer weights from the selection
    Public Sub RemoveAnswerWeightsFromTheSelection()
        With Globals.ThisDocument.Application.Selection.Find
            .ClearFormatting()
            .Style = STYLE_ANSWERWEIGHT
            .Text = ""
            .Replacement.Text = ""
            .Forward = True
            .Format = True
            .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
        End With
    End Sub


    ' Checks the questionnaire.
    ' Returns true if everything is fine, otherwise false
    Function CheckQuestionnaire() As Boolean
        'return false if empty document
        If getDocumentCharacterCount() = 1 Then Return False

        Dim startOfQuestion, endOfQuestion, setEndPoint
        Dim isOK As Boolean

        isOK = True
        setEndPoint = False ' indicates whether the question end point should be set
        startOfQuestion = 0
        questionType = ""

        ' Check each paragraph at a time and specify needed tags
        For Each para As Paragraph In getDocumentParagraphs()

            ' Check if empty paragraph
            If para.Range.Text = vbCr Then
                para.Range.Delete() ' delete all empty paragraphs
                If questionType = "" Then questionType = para.Range.Style.NameLocal
                ' #modif category
            ElseIf para.Range.Style.NameLocal = STYLE_CATEGORYQ Or _
                   para.Range.Style.NameLocal = STYLE_MULTICHOICEQ Or _
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
            ElseIf para.Range.Style.NameLocal = STYLE_CORRECT_MC_ANSWER And _
                   questionType = STYLE_NUMERICALQ Then
                ' Exit if error is found
                If CheckNumericAnswer(para.Range) = False Then Return False 'Exit Function
            End If

            ' Check if the end of document
            If para.Range.End = getDocumentRangeEnd() And _
            startOfQuestion <> para.Range.End Then
                isOK = CheckQuestion(startOfQuestion, getDocumentRangeEnd())
            End If

            'Check if Category Style exist
            If StyleFound(STYLE_CATEGORYQ) = -1 Then
                MsgBox("We must have Ctegory style in your style list")
                isOK = False
            End If


            If isOK = False Then Exit For ' Exit if error is found

        Next para

        'TODO not sure this makes sense, it will just skip the refresh
        If getDocumentCharacterCount() = 1 Then Return isOK


        moveCursorToEndOfDocument()
        Globals.ThisDocument.Application.ScreenRefresh()
        Return isOK
    End Function

    ' Checks whether the chosen question is valid
    ' Returns true if the question is OK, otherwise
    Function CheckQuestion(startPoint As Integer, endPoint As Integer) As Boolean
        Dim isOk As Boolean
        Dim rightCount, rightPairCount, leftPairCount, wordCount, feedbackCount, wrongCount, feedbackTSCount, feedbackFSCount As Integer

        Dim aRange As Range

        aRange = getDocumentRange(startPoint, endPoint)
        aRange.Select()
        'MsgBox "See Range for specifying question type." & questionType & vbCr & _
        '      "Start: " & startPoint & " End: " & endPoint
        isOk = True 'no errors

        ' #modif category
        If questionType = STYLE_CATEGORYQ Then

            rightCount = CountStylesInRange(STYLE_MULTICHOICEQ, startPoint, endPoint)

        ElseIf questionType = STYLE_MULTICHOICEQ Or _
          questionType = STYLE_MULTICHOICEQ_FIXANSWER Then

            ' Check that there are right anwers specified
            rightCount = CountStylesInRange(STYLE_CORRECT_MC_ANSWER, startPoint, endPoint)

            If rightCount = 0 Then
                aRange.Select()
                MsgBox("Error, no correct answer defined.", vbExclamation)
                isOk = False
            End If

            ' Check that there are right feedback specified #modif feedback
            feedbackCount = CountStylesInRange(STYLE_FEEDBACK, startPoint, endPoint)
            wrongCount = CountStylesInRange(STYLE_INCORRECT_MC_ANSWER, startPoint, endPoint)
            If feedbackCount <> rightCount + wrongCount And feedbackCount > 0 Then
                aRange.Select()
                MsgBox("Error, no feedback was supplied for one of answer for this question.", vbExclamation)
                isOk = False
                'ElseIf feedbackCount = 0 Then
                '    isOk = True
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
            rightPairCount = CountStylesInRange(STYLE_RIGHT_MATCH, startPoint, endPoint)
            leftPairCount = CountStylesInRange(STYLE_LEFT_MATCH, startPoint, endPoint)

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
               questionType = STYLE_FALSESTATEMENT Then
            ' Check that there are right feedback specified #modif feedback
            feedbackTSCount = CountStylesInRange(STYLE_FEEDBACK_TS, startPoint, endPoint)
            feedbackFSCount = CountStylesInRange(STYLE_FEEDBACK_FS, startPoint, endPoint)

            If feedbackTSCount = feedbackFSCount = 1 Then
                aRange.Select()
                MsgBox("Error, no correct feedback defined.", vbExclamation)
                isOk = False
            End If

            '  feedbackCount = CountStylesInRange(STYLE_FEEDBACK, startPoint, endPoint)

        ElseIf questionType = STYLE_ESSAY Then
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
    Sub InsertParagraphAfterCurrentParagraph(ByVal text As String, aStyle As String)
        Dim aRange As Range = Globals.ThisDocument.Application.Selection.Paragraphs(1).Range

        With aRange
            .EndOf(Unit:=WdUnits.wdParagraph, Extend:=WdMovementType.wdMove)
            .InsertParagraphBefore()
            .Move(Unit:=WdUnits.wdParagraph, Count:=-1)
            .Style = aStyle
            .InsertBefore(text)
            .Select()
        End With

    End Sub

    ' Inserts text before range found
    Public Sub InsertParagraphOfStyleInSelectedRangeBefore(aStyle, text, index)
        Dim aRange As Microsoft.Office.Interop.Word.Range = getDocumentSelectionRange()
        aRange.Start = index - 1
        With aRange
            .StartOf(Unit:=WdUnits.wdParagraph, Extend:=WdMovementType.wdMove)
            .InsertParagraphBefore()
            .Style = aStyle
            .InsertBefore(text)
            .Select()
        End With
    End Sub
    ' Set the answer weights of multiple choice questions.
    'TODO dead code (never called)
    Public Sub SetAnswerWeights() ' aStyle, startPoint, endPoint)
        Dim startPoint, endPoint, rightScore, wrongScore, rightCount, wrongCount As Integer

        If getSelectionStyleName() = STYLE_MULTICHOICEQ Or STYLE_MULTICHOICEQ_FIXANSWER Then
            startPoint = Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.Start
            rightCount = 0
            wrongCount = 0
            Globals.ThisDocument.Application.Selection.MoveDown(Unit:=WdUnits.wdParagraph, Count:=1)

            Do While getSelectionStyleName() = STYLE_CORRECT_MC_ANSWER Or _
                  getSelectionStyleName() = STYLE_INCORRECT_MC_ANSWER Or _
                  getSelectionStyleName() = STYLE_FEEDBACK Or _
                  getSelectionStyleName() = STYLE_ANSWERWEIGHT

                'Delete empty paragraphs
                If Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.Text = vbCr Then
                    Globals.ThisDocument.Application.Selection.Paragraphs(1).Range.Delete() ' delete all empty paragraphs
                    ' Remove old answer weights
                ElseIf getSelectionStyleName() = STYLE_ANSWERWEIGHT Then
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
                If getSelectionStyleName() = STYLE_CORRECT_MC_ANSWER Then
                    rightCount = rightCount + 1
                ElseIf getSelectionStyleName() = STYLE_INCORRECT_MC_ANSWER Then
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
            ElseIf para.Range.Style = STYLE_CORRECT_MC_ANSWER Then
                InsertAnswerWeight(rightScore, para.Range)
            ElseIf para.Range.Style = STYLE_INCORRECT_MC_ANSWER Then
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



    Public Sub Convert2XML()

        ' Macro recorded on 21.12.2008 by Daniel Refresh Header (translation?)
        Globals.ThisDocument.Application.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekCurrentPageHeader
        Globals.ThisDocument.Application.Selection.Fields.Update()
        Globals.ThisDocument.Application.Selection.EndKey(Unit:=WdUnits.wdLine)
        Globals.ThisDocument.Application.Selection.MoveLeft(Unit:=WdUnits.wdCharacter, Count:=1)
        Globals.ThisDocument.Application.Selection.Fields.Update()
        Globals.ThisDocument.Application.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument


        'choose the file name to save with
        'Dim fd As SaveFileDialog = New SaveFileDialog

        Dim fd As Microsoft.Office.Core.FileDialog = Globals.ThisDocument.Application.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogSaveAs)
        '.FilterIndex = 2 fuer Word 2003, 14 fuer Word 2010
        fd.FilterIndex = 14
        fd.InitialFileName = FILE_PREFIX & Format(Now, "yyyyMMdd") & ".xml"
        If fd.Show <> -1 Then Exit Sub

        Dim header As String
        header = getDocumentHeaderText()


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
        ''write the categories from the header '#modif category
        'objStream.WriteText("<question type=""category"">" & vbCr)
        'objStream.WriteText("<category>" & vbCr)
        'objStream.WriteText("<text>" & header & "</text>" & vbCr)
        'objStream.WriteText("</category>" & vbCr)
        'objStream.WriteText("</question>" & vbCr & vbCr)

        Dim dd As MSXML2.DOMDocument60
        Dim xmlnod As MSXML2.IXMLDOMNode

        'Dim dd As Xml.XmlDocument
        'Dim xmlnod As XMLNode

        'Dim xmlnodelist As MSXML2.IXMLDOMNodeList
        Dim para As Paragraph, paralookahead As Paragraph
        paralookahead = Nothing

        Dim rac, wac As Integer
        Dim xmlResource As String
        Dim i As Integer = 0

        For Each para In getDocumentParagraphs() '?handle each paragraph separately.
            dd = New MSXML2.DOMDocument60

            Select Case para.Range.Style.NameLocal


                Case STYLE_CATEGORYQ
                    'write the categories '#modif category
                    objStream.WriteText("<question type=""category"">" & vbCr)
                    objStream.WriteText("<category>" & vbCr)
                    objStream.WriteText("<text>" & RemoveCR(para.Range.Text) & "</text>" & vbCr)
                    objStream.WriteText("</category>" & vbCr)
                    objStream.WriteText("</question>" & vbCr & vbCr)

                Case STYLE_SHORTANSWERQ
                    xmlResource = My.Resources.Shortanswer_xml
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)
                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)
                    Do While (paralookahead.Style.NameLocal = STYLE_SHORT_ANSWER)
                        xmlnod.attributes.getNamedItem("fraction").text = "100"
                        xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        dd.documentElement.appendChild(xmlnod)

                        xmlnod = xmlnod.cloneNode(True)
                        paralookahead = paralookahead.Next
                        If paralookahead Is Nothing Then Exit Do
                    Loop

                Case STYLE_ESSAY
                    xmlResource = My.Resources.Essay_xml
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)
                    'LTG: do I need to have Set paralookahead = para.Next here? why / why not?

                Case STYLE_NUMERICALQ
                    xmlResource = My.Resources.Numerical_xml
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)
                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)

                    Do While (paralookahead.Style.NameLocal = STYLE_SHORT_ANSWER)
                        xmlnod.attributes.getNamedItem("fraction").text = "100"
                        xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                        paralookahead = paralookahead.Next
                        If Not paralookahead Is Nothing Then
                            If (paralookahead.Style.NameLocal = STYLE_NUM_TOLERANCE) Then
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
                    xmlResource = My.Resources.False_xml
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)
                    paralookahead = para.Next
                    Do While i < 2
                        xmlnod = dd.documentElement.selectSingleNode("answer")
                        If xmlnod.attributes.getNamedItem("fraction").text = "100" Then
                            If paralookahead.Style.NameLocal = STYLE_FEEDBACK_FS Then
                                If paralookahead.Range.Text = "" Then
                                    MsgBox("no feedback was supplied for the false statement")
                                Else
                                    ' Set XML <feedback> text
                                    xmlnod.selectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.Text)
                                    paralookahead = paralookahead.Next
                                End If
                            End If
                            dd.documentElement.appendChild(xmlnod)
                        End If
                        If xmlnod.attributes.getNamedItem("fraction").text = "0" Then
                            'xmlnod.selectSingleNode("text").text = "True"
                            If paralookahead.Style.NameLocal = STYLE_FEEDBACK_TS Then
                                If paralookahead.Range.Text = "" Then
                                    MsgBox("no feedback was supplied for the false statement")
                                Else
                                    ' Set XML <feedback> text
                                    xmlnod.selectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.Text)
                                    paralookahead = paralookahead.Next
                                End If
                            End If
                            dd.documentElement.appendChild(xmlnod)
                        End If
                        i += 1
                    Loop


                Case STYLE_TRUESTATEMENT
                    xmlResource = My.Resources.True_xml
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)
                    paralookahead = para.Next


                    i = 0
                    Do While i < 2
                        xmlnod = dd.documentElement.selectSingleNode("answer")
                        If xmlnod.attributes.getNamedItem("fraction").text = "100" Then
                            If paralookahead.Style.NameLocal = STYLE_FEEDBACK_TS Then
                                If paralookahead.Range.Text = "" Then
                                    MsgBox("no feedback was supplied for the true statement")
                                Else
                                    ' Set XML <feedback> text
                                    xmlnod.selectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.Text)
                                    paralookahead = paralookahead.Next
                                End If
                            End If
                            dd.documentElement.appendChild(xmlnod)
                        End If
                        If xmlnod.attributes.getNamedItem("fraction").text = "0" Then
                            If paralookahead.Style.NameLocal = STYLE_FEEDBACK_FS Then
                                If paralookahead.Range.Text = "" Then
                                    MsgBox("no feedback was supplied for the false statement")
                                Else
                                    ' Set XML <feedback> text
                                    xmlnod.selectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.Text)
                                    paralookahead = paralookahead.Next
                                End If
                            End If
                            dd.documentElement.appendChild(xmlnod)
                        End If
                        i += 1
                    Loop

                Case STYLE_MULTICHOICEQ_FIXANSWER, STYLE_MULTICHOICEQ
                    If para.Range.Style.NameLocal = STYLE_MULTICHOICEQ_FIXANSWER Then
                        xmlResource = My.Resources.MultiChoiceFix_xml
                    Else
                        xmlResource = My.Resources.MultiChoiceVar_xml
                    End If
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)

                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    dd.documentElement.removeChild(xmlnod)
                    rac = 0 'right answer choices
                    wac = 0 'wrong answer choices

                    Do While (paralookahead.Style.NameLocal = STYLE_CORRECT_MC_ANSWER) Or (paralookahead.Style.NameLocal = STYLE_INCORRECT_MC_ANSWER)
                        If paralookahead.Style.NameLocal = STYLE_CORRECT_MC_ANSWER Then
                            xmlnod.attributes.getNamedItem("fraction").text = "100"
                            rac = rac + 1
                        Else
                            xmlnod.attributes.getNamedItem("fraction").text = "0"
                            wac = wac + 1
                        End If
                        ' xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)

                        'insert image in answer
                        'processing <image>'

                        'Create a CData section. 
                        Dim CDATASection As IXMLDOMCDATASection
                        CDATASection = dd.createCDATASection("<p>" & XSLT_Range(paralookahead.Range, My.Resources.FormattedText_xslt) & "<img src=""@@PLUGINFILE@@/image.gif"" width=""88"" height=""74""/></p>")
                        xmlnod.selectSingleNode("text").appendChild(CDATASection)
                        If Not XSLT_Range(paralookahead.Range, My.Resources.PictureName_xslt) = "" Then 'if it is NOT null/empty
                            '   Dim header As String
                            Dim stringlength As Long
                            header = getDocumentHeaderText() ' Globals.ThisDocument.Application.ActiveDocument.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text
                            stringlength = Len(header)
                            header = Left(header, stringlength - 1)
                            'processing <image_base64>'
                            xmlnod.selectSingleNode("file").text = XSLT_Range(paralookahead.Range, My.Resources.Picture_xslt)
                        Else
                            xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                            xmlnod.selectSingleNode("file").text = ""
                        End If
                        ' fin bloc to insert image in answer

                        paralookahead = paralookahead.Next

                        ' Answer Feedback Style processing here
                        If paralookahead IsNot Nothing Then
                            If paralookahead.Style.NameLocal = STYLE_FEEDBACK Then
                                If paralookahead.Range.Text = "" Then
                                    MsgBox("feedback dosn't exist")
                                Else
                                    ' Set XML <feedback> text
                                    xmlnod.selectSingleNode("feedback/text").text = RemoveCR(paralookahead.Range.Text)
                                    paralookahead = paralookahead.Next
                                End If
                            End If
                        End If
                        dd.documentElement.appendChild(xmlnod)
                        xmlnod = xmlnod.cloneNode(True)

                        xmlnod.selectSingleNode("text").text = Nothing
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

                Case STYLE_MATCHINGQ, STYLE_MATCHINGQ_FIXANSWER

                    If para.Range.Style.NameLocal = STYLE_MATCHINGQ Then
                        xmlResource = My.Resources.MatchingVar_xml
                    Else
                        xmlResource = My.Resources.MatchingFix_xml
                    End If
                    loadXML(xmlResource, dd)
                    ProcessCommonTags(dd, para)

                    ' processing each <subquestion>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("subquestion")
                    dd.documentElement.removeChild(xmlnod)
                    Do While (paralookahead.Style.NameLocal = STYLE_LEFT_MATCH)
                        'process left
                        Dim leftQuestion As String = RemoveCR(paralookahead.Range.Text)
                        xmlnod.selectSingleNode("text").text = leftQuestion
                        paralookahead = paralookahead.Next
                        'process right
                        If paralookahead.Style.NameLocal = STYLE_RIGHT_MATCH Then
                            xmlnod.selectSingleNode("answer").selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                            paralookahead = paralookahead.Next
                        Else
                            'error, right is not matching left (should be found in check prior to calling here)
                            Throw New Exception("No matching answer to left question '" & leftQuestion & "'")
                        End If
                        dd.documentElement.appendChild(xmlnod)
                        xmlnod = xmlnod.cloneNode(True)
                        If paralookahead Is Nothing Then Exit Do 'end of questions
                    Loop

                    'Case STYLE_MATCHINGQ_FIXANSWER
                    '    xmlResource = My.Resources.MatchingFix_xml
                    '    loadXML(xmlResource, dd)
                    '    ProcessCommonTags(dd, para)

                    '    ' processing each <subquestion>'
                    '    paralookahead = para.Next
                    '    xmlnod = dd.documentElement.selectSingleNode("subquestion")
                    '    dd.documentElement.removeChild(xmlnod)
                    '    Do While (paralookahead.Style.NameLocal = STYLE_LEFT_PAIR) Or (paralookahead.Style.NameLocal = STYLE_RIGHT_PAIR)
                    '        If paralookahead.Style.NameLocal = STYLE_LEFT_PAIR Then
                    '            xmlnod.selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)
                    '        Else
                    '            xmlnod.selectSingleNode("answer").selectSingleNode("text").text = RemoveCR(paralookahead.Range.Text)

                    '            dd.documentElement.appendChild(xmlnod)
                    '            xmlnod = xmlnod.cloneNode(True)
                    '        End If
                    '        paralookahead = paralookahead.Next
                    '        If paralookahead Is Nothing Then Exit Do
                    '    Loop

                Case STYLE_MISSINGWORDQ
                    'TODO: Verify that MissingWord uses same XML as short answer?
                    xmlResource = My.Resources.Shortanswer_xml
                    loadXML(xmlResource, dd)
                    Dim theChar As Range
                    Dim misword As String
                    misword = ""
                    For Each theChar In para.Range.Characters
                        If theChar.Style.NameLocal = STYLE_BLANK_WORD Then misword = misword & theChar.Text
                    Next theChar
                    ProcessCommonTags(dd, para)  ' XSLT template will swap out missing word
                    dd.documentElement.selectSingleNode("name").selectSingleNode("text").text = _
                        Replace(dd.documentElement.selectSingleNode("name").selectSingleNode("text").text, misword, "__________")

                    ' processing each <answer>'
                    paralookahead = para.Next
                    xmlnod = dd.documentElement.selectSingleNode("answer")
                    xmlnod.attributes.getNamedItem("fraction").text = "100"
                    xmlnod.selectSingleNode("text").text = misword

                    'Case STYLE_COMMENT
                    '    Dim Comment As String
                    '    Comment = "<!-- " & RemoveCR(para.Range.Text) & " -->"
                    '    objStream.WriteText(Comment & vbCr & vbCr)
                    '    dd.loadXML("")

                Case Else
                    dd.loadXML("")
            End Select


            If Not paralookahead Is Nothing Then
                If (paralookahead.Style.NameLocal = STYLE_QUESTIONNAME) Then
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
                '          If (paralookahead.Style.NameLocal = STYLE_FEEDBACK) Then
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
        objStream.SaveToFile(getFileNameFromFileDialog(fd), ADODB.SaveOptionsEnum.adSaveCreateOverWrite)
        'objStream.SaveToFile(FileName:=fd.SelectedItems(1), Options:=ADODB.SaveOptionsEnum.adSaveCreateOverWrite)

    End Sub

    'This is called as the first processing task for each question. It
    Private Sub ProcessCommonTags(dd As MSXML2.DOMDocument60, para As Paragraph)
        ' processing <name> '
        dd.documentElement.selectSingleNode("name") _
        .selectSingleNode("text").text = RemoveCR(para.Range.Text)

        ' processing <questiontext> '
        dd.documentElement.selectSingleNode("questiontext") _
        .selectSingleNode("text").text = XSLT_Range(para.Range, My.Resources.FormattedText_xslt)

        '
        If Not XSLT_Range(para.Range, My.Resources.PictureName_xslt) = "" Then 'if it is NOT null/empty

            Dim header As String
            Dim stringlength As Long

            header = getDocumentHeaderText() ' Globals.ThisDocument.Application.ActiveDocument.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text
            stringlength = Len(header)
            header = Left(header, stringlength - 1)

            'processing <image>'
            dd.documentElement.selectSingleNode("image").text = "Images_forQuizQuestions/" & header & Right(XSLT_Range(para.Range, My.Resources.PictureName_xslt), 4)
            'dd.documentElement.SelectSingleNode("image").text = Mid(XSLT_Range(para.Range, My.Resources.PictureName_xslt), 10) (commented out: Rohrer)'
            'processing <image_base64>'
            dd.documentElement.selectSingleNode("image_base64").text = XSLT_Range(para.Range, My.Resources.Picture_xslt)
        End If

    End Sub


    Private Function XSLT_Range(textrange As Range, xsltFileContents As String) As String
        Dim xsldoc As New MSXML2.FreeThreadedDOMDocument60
        xsldoc.loadXML(xsltFileContents)
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

    Private Function getSelectionStyleName() As String
        If IsNothing(Globals.ThisDocument.Application.Selection.Paragraphs.Style) Then
            Return ""
        Else
            Return CType(Globals.ThisDocument.Application.Selection.Paragraphs.Style, Word.Style).NameLocal
        End If
    End Function

    Private Sub setSelectionParagraphStyle(theStyle As String)
        Globals.ThisDocument.Application.Selection.Paragraphs.Style = theStyle
    End Sub
    Private Function getDocumentSelectionRange() As Range
        Return Globals.ThisDocument.Application.Selection.Range
    End Function

    Private Function getDocumentCharacterCount() As Integer
        Return Globals.ThisDocument.Application.ActiveDocument.Characters.Count
    End Function

    Private Function getDocumentParagraphs() As Paragraphs
        Return Globals.ThisDocument.Application.ActiveDocument.Paragraphs
    End Function

    Private Function getDocumentRangeEnd() As Integer
        Return Globals.ThisDocument.Application.ActiveDocument.Range.End
    End Function

    Public Function getDocumentRange(startPoint As Integer, endPoint As Integer) As Microsoft.Office.Interop.Word.Range
        Return Globals.ThisDocument.Application.ActiveDocument.Range(startPoint, endPoint)
    End Function

    Private Sub moveCursorToEndOfDocument()
        Globals.ThisDocument.Application.Selection.EndKey(WdUnits.wdStory, Nothing)
    End Sub

    Private Sub moveCursorToStartOfDocument()
        Globals.ThisDocument.Application.Selection.HomeKey(WdUnits.wdStory, Nothing)
    End Sub

    Private Function getDocumentHeaderText() As String
        Return Globals.ThisDocument.Application.ActiveDocument.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text()
    End Function

    Private Function getFileNameFromFileDialog(fd As Microsoft.Office.Core.FileDialog) As String
        Return fd.SelectedItems.Item(1)
    End Function

    Private Sub loadXML(xmlResource As String, dd As MSXML2.DOMDocument60)
        If Not dd.loadXML(xmlResource) Then
            MsgBox("Failed to load XML " & xmlResource & " in program.")
            'TODO fail gracefully?
            Throw New Exception
        End If
    End Sub

    Private Sub updateVersionInfo()
        ' Initialize the version number
        If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
            VERSION_INFO = "unknown Network Deployed"

            'This is a ClickOnce Application
            If Not System.Diagnostics.Debugger.IsAttached Then
                VERSION_INFO = System.Deployment.Application _
                    .ApplicationDeployment.CurrentDeployment _
                        .CurrentVersion.ToString()

                'MsgBox("Started " & vbCrLf & " App Version:" _
                '    & My.Application.Info.Version.ToString() & vbCrLf & _
                '    " Published Version " & VERSION_INFO)
            End If
        End If
    End Sub

    Private Function isSelectionNormalStyle() As Boolean
        'Return (Globals.ThisDocument.Application.Selection.Paragraphs(1).Style = Word.WdBuiltinStyle.wdStyleNormal)
        'Return CType(Globals.ThisDocument.Application.Selection.Paragraphs(1).Style, Word.Style).NameLocal = "Normal"
        Dim normalStyle As Style = Globals.ThisDocument.ThisApplication.ActiveDocument.Styles(Word.WdBuiltinStyle.wdStyleNormal)
        Dim selectionStyle As Style = CType(Globals.ThisDocument.Application.Selection.Paragraphs(1).Style, Word.Style)
        Return selectionStyle.NameLocal.Equals(normalStyle.NameLocal) 'http://stackoverflow.com/a/27295771/1168342
    End Function

End Class
