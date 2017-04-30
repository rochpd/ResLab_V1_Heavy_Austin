Imports System
Imports System.Threading
Imports System.Runtime.Remoting.Messaging
Imports Gnostice.PDFOne
Imports Gnostice.PDFOne.PDFPrinter




Public Class class_Printing


    Public Delegate Function AsyncPrint(ByVal pdf As PDFDocument, ByVal docCounter As Integer) As Boolean    ' Delegate for asynchronous printing

    Public Function PrintMe(ByVal pdf As PDFDocument, Optional ByVal Copies As Integer = 1) As Boolean
        'Routine that does the actual printing

        Dim p As New PDFPrinter
        'Dim pg As New PDFPage
        'Dim m As New Gnostice.PDFOne.PageMargins
        'm.measurementUnit = PDFMeasurementUnit.Millimeters
        'm.leftMargin = 20
        'm.topMargin = 10

        Try
            p.LoadDocument(pdf)
            p.Print()

            Return True

        Catch e1 As Exception
            MsgBox("Error printing: " & e1.Message)
            Return False
        End Try

        p.CloseDocument()
        p.Dispose()

    End Function


    Private Function button1_Click(ByVal sender As Object, ByVal e As EventArgs) As Boolean

        'Print' button 'click' event handler.
        'Calls above routine asynchronously.

        Dim aPrint As New AsyncPrint(AddressOf PrintMe)
        Dim aResult(10) As IAsyncResult
        Dim pdf(10) As PDFDocument

        'Code to populate 'files[]' array with names of files omitted

        For i As Integer = 0 To 9
            aResult(i) = aPrint.BeginInvoke(pdf(i), i.ToString(), Nothing, Nothing)
        Next i

        For j As Integer = 0 To 9
            aPrint.EndInvoke(aResult(j))
        Next j

        Return True

    End Function

End Class



