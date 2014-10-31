Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports MoodleQuestions

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestMethod1()
        Dim moodle As MoodleQuestions.MoodleQuestions = New MoodleQuestions.MoodleQuestions()
        moodle.Convert2XML()

    End Sub

End Class