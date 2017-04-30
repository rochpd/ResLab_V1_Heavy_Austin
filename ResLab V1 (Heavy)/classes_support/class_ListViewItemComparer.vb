

Friend Class ListViewItemComparer
    Implements IComparer
    Private col As Integer
    Private _sort As SortOrder = SortOrder.Ascending

    Public Sub New(column As Integer, sort As Windows.Forms.SortOrder)
        col = column
        _sort = sort
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim returnVal As Integer = -1

        If IsDate(x.SubItems(col).Text) And IsDate(y.SubItems(col).Text) Then
            ' parse LV contents back to DateTime value
            Dim dtX As DateTime = DateTime.Parse(CType(x, ListViewItem).SubItems(col).Text)
            Dim dtY As DateTime = DateTime.Parse(CType(y, ListViewItem).SubItems(col).Text)
            ' compare
            returnVal = DateTime.Compare(dtX, dtY)

        ElseIf IsNumeric(x) Then
            ' parse LV contents back to numeric value
            Dim nX As Decimal = Val(x.SubItems(col).Text)
            Dim nY As Decimal = Val(y.SubItems(col).Text)
            ' compare
            returnVal = Decimal.Compare(nX, nY)
        Else
            ' parse LV contents back to string value
            Dim sX As String = x.SubItems(col).Text
            Dim sY As String = y.SubItems(col).Text
            ' compare
            returnVal = String.Compare(sX, sY)
        End If




        If _sort = SortOrder.Descending Then
            returnVal *= -1
        End If
        Return returnVal

    End Function
End Class



