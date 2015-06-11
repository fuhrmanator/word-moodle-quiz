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
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Runtime.InteropServices

Public Class ThisDocument

    ' Allow context-sensitive menu bar
    Event WindowSelectionChange As ApplicationEvents4_WindowSelectionChangeEventHandler
    Dim instance As ApplicationEvents4_Event
    Dim handler As ApplicationEvents4_WindowSelectionChangeEventHandler
    Dim previousSelectionStyle As String

    '<ComVisibleAttribute(False)> _
    'Public Delegate Sub ApplicationEvents4_WindowSelectionChangeEventHandler( _
    '    Sel As Selection _
    ')
    'Usage
    'Dim instance As New ApplicationEvents4_WindowSelectionChangeEventHandler(AddressOf HandleSelectionChange)

    Public Sub HandleSelectionChange(sel As Selection)
        'System.Diagnostics.Debug.WriteLine("caught WindowSelectionChange Event")
        Me.ribbon.Invalidate()
    End Sub

    Public Sub HandleDocumentChange()
        System.Diagnostics.Debug.WriteLine("caught DocumentChange Event")
        Me.ribbon.Invalidate()
    End Sub

    Public ribbon As Office.IRibbonUI
    Public WithEvents pollSelectionChangeTimer As New Timer

    Private Sub ThisDocument_Startup() Handles Me.Startup
        Me.pollSelectionChangeTimer.Enabled = True
        AddHandler pollSelectionChangeTimer.Tick, AddressOf OnTimedEvent
        pollSelectionChangeTimer.Interval = 100

        'AddHandler Globals.ThisDocument.Application.WindowSelectionChange, AddressOf HandleSelectionChange
        'AddHandler Globals.ThisDocument.Application., AddressOf HandleDocumentChange
       

    End Sub

    ' Invalidate the Ribbon to refresh the button states when change selection 
    Public Sub OnTimedEvent(source As Object, e As System.EventArgs)
        Dim currentSelectionParagraphStyle = Globals.ThisDocument.Application.Selection.Paragraphs.Style
        Dim currentSelectionStyle As String
        Dim isSelectionDifferent As Boolean
        If Not IsNothing(currentSelectionParagraphStyle) Then
            currentSelectionStyle = CType(currentSelectionParagraphStyle, Word.Style).NameLocal
            isSelectionDifferent = Not currentSelectionStyle = previousSelectionStyle
            If isSelectionDifferent Then
                previousSelectionStyle = currentSelectionStyle
                If Globals.ThisDocument.Application.MouseAvailable And Me.ribbon IsNot Nothing Then
                    ribbon.Invalidate()
                End If
            End If
        End If
    End Sub

    Private Sub ThisDocument_Shutdown() Handles Me.Shutdown
        ' Globals.ThisDocument.Application.NormalTemplate.Save()
    End Sub

    ' Moodle Quiz Ribbon see http://msdn.microsoft.com/en-us/library/aa942955.aspx for design info
    Protected Overrides Function CreateRibbonExtensibilityObject() As  _
    Microsoft.Office.Core.IRibbonExtensibility
        Return New MoodleQuestions()
    End Function


End Class
