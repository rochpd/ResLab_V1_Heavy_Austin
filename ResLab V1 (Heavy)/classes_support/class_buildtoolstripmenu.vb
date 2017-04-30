


Public Class class_buildtoolstripmenu

    Public Function RftMenu() As ToolStrip

        Dim i As Integer = 0, j As Integer = 0
        Dim ts As New ToolStrip
        ts.Name = "ts_contact"
        Dim ts_btn As ToolStripButton = Nothing
        Dim ts_ddbtn As ToolStripDropDownButton = Nothing
        Dim ts_lbl As ToolStripLabel = Nothing

        Dim tsItem_names() As String = {ts.Name & "_lbl_contacttype", ts.Name & "_btn_rft", ts.Name & "_btn_provs", ts.Name & "_btn_spt", ts.Name & "_btn_cpx", ts.Name & "_btn_hast", ts.Name & "_btn_walk", ts.Name & "_lbl_spacer", ts.Name & "_btn_edit", ts.Name & "_btn_trend"}
        Dim tsItem_type() As String = {"Label", "Button", "DropDownButton", "Button", "Button", "Button", "DropDownButton", "Label", "Button", "Button"}
        Dim tsItem_image() As String = {"", "round_add_32x32", "round_add_32x32", "round_add_32x32", "round_add_32x32", "round_add_32x32", "round_add_32x32", "", "edit_32x32", "line_graph_32x32"}
        Dim tsItem_tooltip() As String = {"", "New RFT", "New bronchoprovocation test", "New skin prick test", "New cardio-pulmonary exercise test", "New altitude simulation test", "New walk test", "", "Edit test", "Display trend view"}
        Dim tsItem_text() As String = {"Lung function tests", "RFT", "BHR", "SPT", "CPET", "HAST", "WALK", "", "Edit test", "Display trend"}
        Dim tsItem_sizeX() As Integer = {200, 53, 67, 53, 55, 66, 75, 50, 80, 100}
        Dim tsItem_enabled() As Boolean = {True, True, True, True, True, True, True, False, True, True}

        For i = 0 To UBound(tsItem_names)
            Select Case tsItem_type(i)
                Case "Button"
                    ts_btn = New ToolStripButton
                    ts_btn.Name = tsItem_names(i)
                    ts_btn.Image = CType(My.Resources.ResourceManager.GetObject(tsItem_image(i)), System.Drawing.Image)
                    ts_btn.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_btn.Text = tsItem_text(i)
                    ts_btn.ToolTipText = tsItem_tooltip(i)
                    ts_btn.AutoSize = False
                    ts_btn.Enabled = tsItem_enabled(i)

                    AddHandler ts_btn.Click, AddressOf Form_MainNew.EventHandler_ToolBar_rft

                    ts.Items.Add(ts_btn)

                Case "DropDownButton"
                    ts_ddbtn = New ToolStripDropDownButton
                    ts_ddbtn.Name = tsItem_names(i)
                    ts_ddbtn.Image = CType(My.Resources.ResourceManager.GetObject(tsItem_image(i)), System.Drawing.Image)
                    ts_ddbtn.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_ddbtn.Text = tsItem_text(i)
                    ts_ddbtn.ToolTipText = tsItem_tooltip(i)
                    ts_ddbtn.AutoSize = False
                    ts_ddbtn.Enabled = tsItem_enabled(i)

                    'Add subitems
                    Select Case ts_ddbtn.Name
                        Case ts.Name & "_btn_provs" : Me.Add_ProvProtocols(ts_ddbtn)
                        Case ts.Name & "_btn_walk" : Me.Add_WalkProtocols(ts_ddbtn)
                    End Select

                    For j = 0 To ts_ddbtn.DropDownItems.Count - 1
                        AddHandler ts_ddbtn.DropDownItems(j).Click, AddressOf Form_MainNew.EventHandler_ToolBar_rft
                    Next

                    ts.Items.Add(ts_ddbtn)

                Case "Label"
                    ts_lbl = New ToolStripLabel
                    ts_lbl.Name = tsItem_names(i)
                    ts_lbl.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_lbl.Text = tsItem_text(i)
                    ts_lbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                    ts_lbl.AutoSize = False
                    ts_lbl.Enabled = tsItem_enabled(i)

                    ts.Items.Add(ts_lbl)

            End Select


        Next
        
        Me.Add_GenericButtonsToToolstrip(ts)
        ts.BackColor = Color.FromName("GradientInactiveCaption")
        ts.Dock = DockStyle.Top

        Return ts

    End Function

    Private Sub Add_ProvProtocols(ByRef btn As ToolStripDropDownButton)

        Dim p() As class_challenge.ProtocolMenuData = cChall.Get_Protocols(True)
        Dim mi As ToolStripMenuItem = Nothing

        If Not IsNothing(p) Then
            For i As Integer = 0 To UBound(p)
                mi = New ToolStripMenuItem
                mi.Name = btn.Name & "_" & p(i).Label
                mi.Text = p(i).Label
                mi.Tag = p(i).protocolID.ToString
                btn.DropDownItems.Add(mi)
            Next
        End If

    End Sub

    Private Sub Add_WalkProtocols(ByRef btn As ToolStripDropDownButton)

        Dim p() As class_walktest.ProtocolMenuData = cWalk.Get_Protocols(True)
        Dim mi As ToolStripMenuItem = Nothing

        If Not IsNothing(p) Then
            For i As Integer = 0 To UBound(p)
                mi = New ToolStripMenuItem
                mi.Name = btn.Name & "_" & p(i).Label
                mi.Text = p(i).Label
                mi.Tag = p(i).protocolID.ToString
                btn.DropDownItems.Add(mi)
            Next
        End If

    End Sub

    Private Sub Add_GenericButtonsToToolstrip(ByRef ts As ToolStrip)

        Dim i As Integer = 0, j As Integer = 0

        Dim ts_btn As ToolStripButton = Nothing
        Dim ts_ddbtn As ToolStripDropDownButton = Nothing
        Dim ts_lbl As ToolStripLabel = Nothing

        Dim mi As ToolStripMenuItem = Nothing
        Dim mi_names() As String = Nothing
        Dim mi_text() As String = Nothing
        Dim mi_tag() As String = Nothing

        Dim tsItem_names() As String = {ts.Name & "_btn_savetofile", ts.Name & "_btn_zoom", ts.Name & "_btn_print"}
        Dim tsItem_type() As String = {"Button", "DropDownButton", "DropDownButton"}
        Dim tsItem_image() As String = {"save_32x32", "zoom_in_32x32", "print_32x32"}
        Dim tsItem_tooltip() As String = {"Save PDF to file", "PDF view options", "Print PDF"}
        Dim tsItem_text() As String = {"SAVETOFILE", "ZOOM", "PRINT"}
        Dim tsItem_sizeX() As Integer = {23, 29, 29}
        Dim tsItem_enabled() As Boolean = {True, True, True}

        For i = 0 To UBound(tsItem_names)
            Select Case tsItem_type(i)
                Case "Button"
                    ts_btn = New ToolStripButton
                    ts_btn.Name = tsItem_names(i)
                    ts_btn.Image = CType(My.Resources.ResourceManager.GetObject(tsItem_image(i)), System.Drawing.Image)
                    ts_btn.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_btn.Text = tsItem_text(i)
                    ts_btn.ToolTipText = tsItem_tooltip(i)
                    ts_btn.AutoSize = False
                    ts_btn.Alignment = ToolStripItemAlignment.Right
                    ts_btn.Enabled = tsItem_enabled(i)
                    ts_btn.DisplayStyle = ToolStripItemDisplayStyle.Image
                    AddHandler ts_btn.Click, AddressOf Form_MainNew.EventHandler_ToolBar_generic
                    ts.Items.Add(ts_btn)

                Case "DropDownButton"
                    ts_ddbtn = New ToolStripDropDownButton
                    ts_ddbtn.Name = tsItem_names(i)
                    ts_ddbtn.Image = CType(My.Resources.ResourceManager.GetObject(tsItem_image(i)), System.Drawing.Image)
                    ts_ddbtn.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_ddbtn.Text = tsItem_text(i)
                    ts_ddbtn.ToolTipText = tsItem_tooltip(i)
                    ts_ddbtn.AutoSize = False
                    ts_ddbtn.Alignment = ToolStripItemAlignment.Right
                    ts_ddbtn.Enabled = tsItem_enabled(i)
                    ts_ddbtn.DisplayStyle = ToolStripItemDisplayStyle.Image

                    'Add subitems                   
                    Select Case ts_ddbtn.Name
                        Case ts.Name & "_btn_zoom"
                            mi_names = {ts.Name & "_fitwidth", ts.Name & "_fitpage", ts.Name & "_clearpage"}
                            mi_text = {"Fit width", "Fit page", "Clear page"}
                            mi_tag = {"", "", ""}
                        Case ts.Name & "_btn_print"
                            mi_names = {ts.Name & "_report", ts.Name & "_trend"}
                            mi_text = {"Print report page", "Print trend page"}
                            mi_tag = {"", "", ""}
                    End Select
                    For j = 0 To UBound(mi_names)
                        mi = New ToolStripMenuItem
                        mi.Name = ts_ddbtn.Name & "_" & mi_names(j)
                        mi.Text = mi_text(j)
                        mi.Tag = mi_tag(j)
                        ts_ddbtn.DropDownItems.Add(mi)
                    Next
                    'If mi.Name = "ts_contact_btn_print_ts_contact_trend" Then mi.Enabled = False

                    For j = 0 To ts_ddbtn.DropDownItems.Count - 1
                        AddHandler ts_ddbtn.DropDownItems(j).Click, AddressOf Form_MainNew.EventHandler_ToolBar_generic
                    Next

                    ts.Items.Add(ts_ddbtn)
                Case "Label"
                    ts_lbl = New ToolStripLabel
                    ts_lbl.Name = tsItem_names(i)
                    ts_lbl.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_lbl.Text = tsItem_text(i)
                    ts_lbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                    ts_lbl.AutoSize = False
                    ts_lbl.Alignment = ToolStripItemAlignment.Right
                    ts_lbl.Enabled = tsItem_enabled(i)
                    ts.Items.Add(ts_lbl)
            End Select
        Next

    End Sub

    Public Function SleepMenu() As ToolStrip

        Dim i As Integer = 0, j As Integer = 0
        Dim ts As New ToolStrip
        ts.Name = "ts_contact"
        Dim ts_btn As ToolStripButton = Nothing
        Dim ts_ddbtn As ToolStripDropDownButton = Nothing
        Dim ts_lbl As ToolStripLabel = Nothing

        Dim tsItem_names() As String = {ts.Name & "_lbl_contacttype", ts.Name & "_btn_psg", ts.Name & "_lbl_spacer", ts.Name & "_btn_edit"}
        Dim tsItem_type() As String = {"Label", "Button", "Label", "Button"}
        Dim tsItem_image() As String = {"", "round_add_32x32", "", "edit_32x32"}
        Dim tsItem_tooltip() As String = {"", "New PSG", "", "Edit RFT"}
        Dim tsItem_text() As String = {"Sleep studies", "PSG", "", "Edit"}
        Dim tsItem_sizeX() As Integer = {200, 53, 67, 53}
        Dim tsItem_enabled() As Boolean = {True, True, False, True}

        For i = 0 To UBound(tsItem_names)
            Select Case tsItem_type(i)
                Case "Button"
                    ts_btn = New ToolStripButton
                    ts_btn.Name = tsItem_names(i)
                    ts_btn.Image = CType(My.Resources.ResourceManager.GetObject(tsItem_image(i)), System.Drawing.Image)
                    ts_btn.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_btn.Text = tsItem_text(i)
                    ts_btn.ToolTipText = tsItem_tooltip(i)
                    ts_btn.AutoSize = False
                    ts_btn.Enabled = tsItem_enabled(i)

                    AddHandler ts_btn.Click, AddressOf Form_MainNew.EventHandler_ToolBar_sleep

                    ts.Items.Add(ts_btn)

                Case "DropDownButton"
                    ts_ddbtn = New ToolStripDropDownButton
                    ts_ddbtn.Name = tsItem_names(i)
                    ts_ddbtn.Image = CType(My.Resources.ResourceManager.GetObject(tsItem_image(i)), System.Drawing.Image)
                    ts_ddbtn.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_ddbtn.Text = tsItem_text(i)
                    ts_ddbtn.ToolTipText = tsItem_tooltip(i)
                    ts_ddbtn.AutoSize = False
                    ts_ddbtn.Enabled = tsItem_enabled(i)

                    'Add subitems
                    Select Case ts_ddbtn.Name
                        Case ts.Name & "_"
                    End Select

                    For j = 0 To ts_ddbtn.DropDownItems.Count - 1
                        AddHandler ts_ddbtn.DropDownItems(j).Click, AddressOf Form_MainNew.EventHandler_ToolBar_sleep
                    Next

                    ts.Items.Add(ts_ddbtn)

                Case "Label"
                    ts_lbl = New ToolStripLabel
                    ts_lbl.Name = tsItem_names(i)
                    ts_lbl.Size = New System.Drawing.Size(tsItem_sizeX(i), 29)
                    ts_lbl.Text = tsItem_text(i)
                    ts_lbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                    ts_lbl.AutoSize = False
                    ts_lbl.Enabled = tsItem_enabled(i)

                    ts.Items.Add(ts_lbl)

            End Select


        Next

        Me.Add_GenericButtonsToToolstrip(ts)
        ts.BackColor = Color.FromName("GradientInactiveCaption")
        ts.Dock = DockStyle.Top

        Return ts

    End Function
  
End Class
