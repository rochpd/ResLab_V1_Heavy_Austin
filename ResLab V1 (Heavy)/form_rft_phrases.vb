
Public Class form_rft_phrases

    Private _reportTextBox As RichTextBox
    Private TextForUndo As String = ""
    Private _demo As Pred_demo
    Private _r As Dictionary(Of String, String)
    Private _eTestGroup As eAutoreport_testgroups
    Private _callingForm As Form

    Public Property ReportTextBox() As RichTextBox
        Get
            Return _reportTextBox
        End Get
        Set(ByVal Value As RichTextBox)
            _reportTextBox = Value
        End Set
    End Property

    Public Sub New(eTestGroup As eAutoreport_testgroups, CallingForm As Form, Optional Demo As Pred_demo = Nothing, Optional R As Dictionary(Of String, String) = Nothing)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _eTestGroup = eTestGroup
        _r = R
        _demo = Demo
        _callingForm = CallingForm


        Me.Top = CallingForm.Top + 50
        Me.Left = CallingForm.Left + CallingForm.Width - Me.Width - 50

    End Sub

    Private Sub form_rft_phrases_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim d As Dictionary(Of Integer, String) = cPhrases.get_phrasegroups(_eTestGroup)
        Dim i As Integer = 0

        lstGroups.Items.Clear()

        For Each kv As KeyValuePair(Of Integer, String) In d
            lstGroups.Items.Add(kv)
        Next
        lstGroups.DisplayMember = "value"
        lstGroups.ValueMember = "key"
        If lstGroups.Items.Count > 0 Then lstGroups.SelectedIndex = 0

        lstGroups.Height = split.Height - ToolStrip2.Height
        lstItems.Height = lstGroups.Height

    End Sub

    Private Sub form_rft_phrases_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize

        lstGroups.Height = split.Height - ToolStrip2.Height
        lstItems.Height = lstGroups.Height

    End Sub

    Private Sub lstGroups_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstGroups.SelectedIndexChanged

        Me.Load_ListItems(CType(lstGroups.SelectedItem, KeyValuePair(Of Integer, String)).Key)

    End Sub

    Private Sub tsbtn_clear_Click(sender As System.Object, e As System.EventArgs) Handles tsbtn_clear.Click

        TextForUndo = _reportTextBox.Text
        _reportTextBox.Text = ""

    End Sub

    Private Sub tsbtn_Undo_Click(sender As System.Object, e As System.EventArgs) Handles tsbtn_Undo.Click

        If TextForUndo <> "" Then
            _reportTextBox.Text = TextForUndo
            TextForUndo = ""
        End If

    End Sub

    Private Sub EditItemToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles EditItemToolStripMenuItem.Click

        If lstItems.SelectedIndex = -1 Then
            MsgBox("Select an entry to edit", vbOKOnly, "Edit list item")
        Else
            Me.TopLevel = False
            Dim txt As String = Nothing
            Dim kv As KeyValuePair(Of Integer, String) = lstItems.Items.Item(lstItems.SelectedIndex)
            txt = InputBox("Edit the text", "Phrase list", kv.Value)
            If txt <> "" Then
                Dim kv1 As New KeyValuePair(Of Integer, String)(kv.Key, txt)
                lstItems.Items.Item(lstItems.SelectedIndex) = kv1
                'Save to database
                cPhrases.Save_phrase(CType(lstGroups.SelectedItem, KeyValuePair(Of Integer, String)).Key, kv1.Key, kv1.Value)
            End If
            Me.TopLevel = True
        End If

    End Sub

    Private Sub lstItems_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstItems.MouseDoubleClick

        Select Case e.Button
            Case Windows.Forms.MouseButtons.Left
                Dim s As String = CType(lstItems.SelectedItem, KeyValuePair(Of Integer, String)).Value
                TextForUndo = _reportTextBox.Text

                If _reportTextBox.Text = "" Or s = "" Then
                    'leave alone
                Else
                    Dim LastChar As String = Strings.Right(_reportTextBox.Text, 1)
                    Dim FirstChar As String = Strings.Left(s, 1)
                    Select Case LastChar
                        Case ","
                            If FirstChar <> " " Then s = " " & s
                        Case "."
                            If FirstChar <> " " Then s = " " & Strings.Left(s, 1).ToUpper & Strings.Mid(s, 2)
                        Case " "
                            'leave alone
                        Case Else
                            If FirstChar <> " " Then s = " " & s
                    End Select
                End If
                _reportTextBox.Text = _reportTextBox.Text & s

            Case Windows.Forms.MouseButtons.Right
                'not used
        End Select

    End Sub

    Private Sub NewItemToolStripMenuItem_Click(sender As Object, e As System.EventArgs) Handles NewItemToolStripMenuItem.Click

        Me.TopLevel = False
        Dim txt As String = Nothing
        txt = InputBox("Edit the text", "Phrase list", "")
        If txt <> "" Then
            Dim NewItemID As Integer = cPhrases.Save_phrase(CType(lstGroups.SelectedItem, KeyValuePair(Of Integer, String)).Key, 0, txt)
            If NewItemID > 0 Then
                Dim kv As New KeyValuePair(Of Integer, String)(NewItemID, txt)
                lstItems.Items.Add(kv)
            Else
                MsgBox("Problem saving new list item", vbOKOnly, "New list item")
            End If
        End If
        Me.TopLevel = True

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As System.EventArgs) Handles DeleteToolStripMenuItem.Click

        cPhrases.Delete_phrase(CType(lstItems.SelectedItem, KeyValuePair(Of Integer, String)).Key)
        lstItems.SelectedItems.Remove(lstItems.SelectedItem)
        Load_ListItems(CType(lstGroups.SelectedItem, KeyValuePair(Of Integer, String)).Key)

    End Sub

    Private Sub Load_ListItems(GroupID As Integer)

        Dim d As Dictionary(Of Integer, String) = cPhrases.get_phrases(GroupID)
        Dim i As Integer = 0

        lstItems.Items.Clear()
        lstItems.ValueMember = "key"
        lstItems.DisplayMember = "value"

        If Not IsNothing(d) Then
            For Each kv As KeyValuePair(Of Integer, String) In d
                lstItems.Items.Add(kv)
            Next
        End If

    End Sub

    Private Sub tsbtnAutoReport_Click(sender As Object, e As System.EventArgs) Handles tsbtnAutoReport.Click

        Dim Autoreport_text As String = cPhrases.AutoInterpret_rft(_eTestGroup, _demo, _r)

        If Trim(_reportTextBox.Text) = "" Then
            _reportTextBox.Text = Autoreport_text
        Else

            Dim form1 As New Form()
            Dim lbl As New Label
            Dim button1 As New Button()
            Dim button2 As New Button()
            Dim button3 As New Button()
            lbl.Text = "Report text already present. Choose between ...."
            lbl.Width = 400
            lbl.TextAlign = ContentAlignment.MiddleCenter
            lbl.Location = New Point(0, 15)
            button1.Text = "Append"
            button1.Location = New Point(70, lbl.Top + lbl.Height + 15)
            button1.DialogResult = DialogResult.OK
            button1.AutoSize = True
            button2.Text = "Overwrite"
            button2.Location = New Point(button1.Left + button1.Width + 5, button1.Top)
            button2.DialogResult = DialogResult.Yes
            button2.AutoSize = True
            button3.Text = "Cancel"
            button3.Location = New Point(button2.Left + button2.Width + 5, button1.Top)
            button3.DialogResult = DialogResult.Cancel
            button3.AutoSize = True

            form1.Font = form_rft_phrases.DefaultFont
            form1.Text = "Auto report"
            form1.ControlBox = False
            form1.Height = 150
            form1.Width = lbl.Width
            form1.FormBorderStyle = FormBorderStyle.FixedSingle
            form1.AcceptButton = button1
            form1.CancelButton = button3
            form1.StartPosition = FormStartPosition.CenterScreen
            form1.Controls.Add(button1)
            form1.Controls.Add(button2)
            form1.Controls.Add(button3)
            form1.Controls.Add(lbl)

            Me.TopMost = False
            form1.ShowDialog()
            Me.TopMost = True

            Select Case form1.DialogResult
                Case DialogResult.OK : _reportTextBox.Text = _reportTextBox.Text & vbCrLf & Autoreport_text
                Case DialogResult.Yes : _reportTextBox.Text = Autoreport_text
                Case DialogResult.Cancel
            End Select
            form1.Dispose()
        End If
    End Sub

    Private Sub lstItems_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles lstItems.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim i As Integer = lstItems.IndexFromPoint(New Point(e.X, e.Y))
            If i >= 0 And i <= lstItems.Items.Count - 1 Then lstItems.SelectedIndex = i
        End If
    End Sub

    
    Private Sub lstItems_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lstItems.SelectedIndexChanged

    End Sub
End Class