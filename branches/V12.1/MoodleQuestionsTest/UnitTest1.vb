Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports MoodleQuestions
Imports Microsoft.Office.Interop.Word
Imports MSXML2

<TestClass()>
Public Class UnitTest1
    <TestMethod()> Public Sub TestMethod1()
        ' Dim aRange As Microsoft.Office.Interop.Word.Range = getDocumentSelectionRange()

        Dim wordApp As Application
        'wordApp = New Application
        Globals.ThisDocument.Application()
        Dim moodle As MoodleQuestions.MoodleQuestions = New MoodleQuestions.MoodleQuestions
        Dim doc As MoodleQuestions.ThisDocument = New MoodleQuestions.ThisDocument
        'Dim doc As MoodleQuestions.a
        ' moodle.Ribbon_Load()
        moodle.app()
        '  moodle.OnLoadImage()
        'class thisDocument.vb?
        moodle.Convert2XML()
        moodle.GetCustomUI(Microsoft.Word.Document)


    End Sub

End Class