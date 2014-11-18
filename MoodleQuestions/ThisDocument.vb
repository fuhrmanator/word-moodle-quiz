'MS WORD Document for making Moodle questions
'=================================================
'(uses Moodle's XML interchange format)
'
'CREDITS
'based on Moodle_Quiz_V21 by Lael Grant (?), previous versions by Vyatcheslav Yatskovsky (yatskovsky@gmail.com) and others
'based on the GIFTconverter template by Mikko Rusama
'inspired by OpenOffice aesthetical version by Enrique Castro
'************************
'The MIT License
'Copyright (c) 2005 Mikko Rusama, 2006 Vyatcheslav Yatskovky
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'************************
Imports Microsoft.Office.Interop.Word
Imports MSXML2
Imports Microsoft.VisualBasic.ApplicationServices

Public Class ThisDocument

    Public ribbon As Office.IRibbonUI
    Public WithEvents timer1 As New Timer

    Private Sub ThisDocument_Startup() Handles Me.Startup
        Me.timer1.Enabled = True
        AddHandler timer1.Tick, AddressOf OnTimedEvent
        timer1.Interval = 100
    End Sub

    ' Invalidate the Ribbon to refresh the button states when change selection 
    Public Sub OnTimedEvent(source As Object, e As System.EventArgs)
        If Globals.ThisDocument.Application.MouseAvailable And Me.ribbon IsNot Nothing Then
            ribbon.Invalidate()
        End If
    End Sub

    Private Sub ThisDocument_Shutdown() Handles Me.Shutdown

    End Sub

    ' Moodle Quiz Ribbon see http://msdn.microsoft.com/en-us/library/aa942955.aspx for design info
    Protected Overrides Function CreateRibbonExtensibilityObject() As  _
    Microsoft.Office.Core.IRibbonExtensibility
        Return New MoodleQuestions()
    End Function


End Class
