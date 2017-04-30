
Imports Gnostice.PDFOne

Public Class class_reportstrings

    Public Structure rep_items
        Public name As String
        Public text As String
        Public pref_displaytext As String
        Public fontname As String
        Public font As PDFFont
        Public fontsize As Integer
        Public fontbold As Boolean
        Public fontitalic As Boolean
    End Structure

    Public ReadOnly rft_reporttitle_routine As rep_items
    Public ReadOnly rft_reporttitle_cpet As rep_items
    Public ReadOnly rft_reporttitle_hast As rep_items
    Public ReadOnly rft_reporttitle_spt As rep_items
    Public ReadOnly rft_reporttitle_bhr As rep_items
    Public ReadOnly rft_reporttitle_walk As rep_items
    Public ReadOnly rft_serviceline_1 As rep_items
    Public ReadOnly rft_serviceline_2 As rep_items
    Public ReadOnly rft_serviceline_3 As rep_items
    Public ReadOnly rft_serviceline_4 As rep_items
    Public ReadOnly rft_serviceline_5 As rep_items
    Public ReadOnly rft_serviceline_6 As rep_items

    'Dim fCourier As PDFFont = New PDFFont(StdType1Font.Courier, PDFFontStyle.Fill, 14)
    Dim fHel8 As PDFFont = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 8)
    Dim fHel8b As PDFFont = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 8)
    Dim fHel10 As PDFFont = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 10)
    Dim fHel10b As PDFFont = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 10)
    Dim fHel10i As PDFFont = New PDFFont(StdType1Font.Helvetica_Oblique, PDFFontStyle.Fill, 10)
    Dim fHel11 As PDFFont = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 11)
    Dim fHel11b As PDFFont = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 11)
    Dim fHel12 As PDFFont = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 12)
    Dim fHel12b As PDFFont = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 12)
    Dim fHel14 As PDFFont = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 14)
    Dim fHel14b As PDFFont = New PDFFont(StdType1Font.Helvetica_Bold, PDFFontStyle.Fill, 16)
    Dim fHel16 As PDFFont = New PDFFont(StdType1Font.Helvetica, PDFFontStyle.Fill, 16)
    'Dim fTimes As PDFFont = New PDFFont(StdType1Font.Times_Roman, PDFFontStyle.Fill, 14)

    Private Function Get_pdfFont(ByVal fontname As String, ByVal fontsize As Integer, ByVal fontbold As Boolean, ByVal fontitalic As Boolean) As PDFFont

        Select Case fontname
            Case "Helvetica"
                If fontbold Then
                    If fontitalic Then
                        Select Case fontsize
                            Case 8 : Return Nothing
                            Case 10 : Return Nothing
                            Case 12 : Return Nothing
                            Case 14 : Return Nothing
                            Case 16 : Return Nothing
                            Case Else : Return Nothing
                        End Select
                    Else
                        Select Case fontsize
                            Case 8 : Return fHel8b
                            Case 10 : Return fHel10b
                            Case 12 : Return fHel12b
                            Case 14 : Return fHel14b
                            Case 16 : Return Nothing
                            Case Else : Return Nothing
                        End Select
                    End If
                Else
                    If fontitalic Then
                        Select Case fontsize
                            Case 8 : Return Nothing
                            Case 10 : Return fHel10i
                            Case 12 : Return Nothing
                            Case 14 : Return Nothing
                            Case 16 : Return Nothing
                            Case Else : Return Nothing
                        End Select
                    Else
                        Select Case fontsize
                            Case 8 : Return fHel8
                            Case 10 : Return fHel10
                            Case 12 : Return fHel12
                            Case 14 : Return fHel14
                            Case 16 : Return fHel16
                            Case Else : Return Nothing
                        End Select
                    End If
                End If
            Case Else
                Return Nothing
        End Select

    End Function

    Public Sub New(Optional ByVal reporttypeID As Integer = 0)


        Dim sql As String = "SELECT * FROM prefs_reports_strings"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)

        If Ds Is Nothing Then

        Else
            Dim item As rep_items
            For Each row As DataRow In Ds.Tables(0).Rows
                item = New rep_items
                item.name = row("Name") & ""
                item.pref_displaytext = row("pref_displaytext") & ""
                item.text = row("text") & ""
                item.fontname = row("font") & ""
                item.fontbold = Convert.ToBoolean(row("font_bold"))
                item.fontsize = row("font_size")
                item.fontitalic = Convert.ToBoolean(row("font_italic"))
                item.font = Get_pdfFont(item.fontname, item.fontsize, item.fontbold, item.fontitalic)

                Select Case row("Name")
                    Case "rft_reporttitle_routine" : Me.rft_reporttitle_routine = item
                    Case "rft_reporttitle_cpet" : Me.rft_reporttitle_cpet = item
                    Case "rft_reporttitle_hast" : Me.rft_reporttitle_hast = item
                    Case "rft_reporttitle_walk" : Me.rft_reporttitle_walk = item
                    Case "rft_reporttitle_spt" : Me.rft_reporttitle_spt = item
                    Case "rft_reporttitle_bhr" : Me.rft_reporttitle_bhr = item
                    Case "rft_serviceline_1" : Me.rft_serviceline_1 = item
                    Case "rft_serviceline_2" : Me.rft_serviceline_2 = item
                    Case "rft_serviceline_3" : Me.rft_serviceline_3 = item
                    Case "rft_serviceline_4" : Me.rft_serviceline_4 = item
                    Case "rft_serviceline_5" : Me.rft_serviceline_5 = item
                    Case "rft_serviceline_6" : Me.rft_serviceline_6 = item
                End Select
            Next
        End If

    End Sub


End Class
