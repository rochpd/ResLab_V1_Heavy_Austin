
Imports System.Net

Public Module module_Initialise

   


    Public Sub InitialiseApp()

        Form_MainNew.ToolStripLabel_UR.Text = gURlabel

        cAppConfig.intialise_app_config()

        'Label main form header with lab name
        Dim rS As New class_reportstrings
        Form_MainNew.Text = "ResLab (Heavy)   -  " & rS.rft_serviceline_1.text
        rS = Nothing

        'Load list of enabled units and set default
        Form_MainNew.Setup_lvUnits()
        Dim lvi As ListViewItem
        Dim dic As Dictionary(Of String, Integer) = cMyRoutines.Get_ListOfUnits(True, True)
        If Not (dic Is Nothing) Then
            For Each kv As KeyValuePair(Of String, Integer) In dic
                lvi = New ListViewItem
                lvi.SubItems.Add(kv.Value)
                lvi.SubItems.Add(kv.Key)
                Form_MainNew.lvUnits.Items.Add(lvi)
            Next
            Form_MainNew.lvUnits.Items(0).Selected = True
            Form_MainNew.lvUnits.Focus()
        End If

    End Sub

    Public Function CheckIfCurrentversion() As Integer
        'returns
        '-1=can't connect to website
        '0=current version, current build
        '1=current version, old build
        '2=old version

        Dim c As Integer = -1
        Dim client As WebClient = New WebClient()

        Try
            Dim version As String = client.DownloadString("http://www.helmutdesign.com.au/reslab/version")
            Dim build As String = client.DownloadString("http://www.helmutdesign.com.au/reslab/build")

            If version = My.Settings.Version Then
                If build = My.Settings.BuildDate Then
                    Return 0
                Else
                    Return 1
                End If
            Else
                Return 2
            End If
        Catch
            Return -1
        End Try

    End Function

    Public Function Get_Currentversion() As String

        Dim client As WebClient = New WebClient()

        Try
            Dim version As String = client.DownloadString("http://www.helmutdesign.com.au/reslab/version")
            Return version
        Catch
            Return ""
        End Try

    End Function

    Public Function Get_CurrentBuild() As String

        Dim client As WebClient = New WebClient()

        Try
            Dim build As String = client.DownloadString("http://www.helmutdesign.com.au/reslab/build")
            Return build
        Catch
            Return ""
        End Try

    End Function

End Module
