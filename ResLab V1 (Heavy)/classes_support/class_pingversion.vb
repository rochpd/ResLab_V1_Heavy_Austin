Imports System.Net
Imports ResLab_V1_Heavy.cDatabaseInfo

Public Class class_pingversion

    Public Sub ping_version()

        Dim labid As String = Me.get_labID

        Dim installation As String = Me.get_installation
        Dim version As String = My.Settings.Version
        Dim build As String = My.Settings.BuildDate

        Dim uri As String = "https://reslab.com.au/api/ping/" & labid & "/" & installation & "/" & version

        Try
            Dim Request As HttpWebRequest = WebRequest.Create(uri)
            Dim Response As HttpWebResponse = Request.GetResponse
            'MsgBox(Response.StatusDescription)
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try


    End Sub

    Public Function get_labID() As String

        Dim sql As String = "SELECT lab_id FROM lab_id WHERE id=1"
        Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
        If Not IsNothing(Ds) Then
            If Ds.Tables(0).Rows.Count = 1 Then
                Return Ds.Tables(0).Rows(0)("lab_id")
            Else
                Return ""
            End If
        Else
            Return ""
        End If

    End Function

    Public Function get_installation() As String
        'Gets the mac address of this pc

        Dim networkInterface = System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
        Dim firstNetwork = networkInterface.FirstOrDefault(Function(x As System.Net.NetworkInformation.NetworkInterface) _
            x.OperationalStatus = System.Net.NetworkInformation.OperationalStatus.Up)
        Dim firstMacAddressOfWorkingNetworkAdapter = firstNetwork.GetPhysicalAddress()

        Return firstMacAddressOfWorkingNetworkAdapter.ToString

    End Function

    Public Function update_labID(labID As String) As Integer
        'Should only be one record in table, or none

        Dim sql As String = ""
        Dim d As New Dictionary(Of String, String)

        Try
            'No record yet?
            sql = "SELECT count(lab_id) FROM lab_id"
            Dim Ds As DataSet = cDAL.Get_DataAsDataset(sql)
            If Not IsNothing(Ds) Then
                d.Add("lab_id", labID)
                d.Add("updated_date", Now)
                Select Case Ds.Tables(0).Rows(0).Item(0)
                    Case 0
                        'Apply insert
                        Try
                            d.Add("id", "0")
                            sql = cDAL.Build_InsertQuery(eTables.site_id, d)
                            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
                            Return ReturnValue
                        Catch ex As Exception
                            MsgBox("Error inserting lab_id record" & vbNewLine & ex.Message.ToString)
                            Return 0
                        End Try
                    Case 1
                        'Update record
                        Try
                            d.Add("id", "1")
                            sql = cDAL.Build_UpdateQuery(eTables.site_id, d)
                            Dim ReturnValue As Long = cDAL.Insert_Record(sql)
                            Return ReturnValue
                        Catch ex As Exception
                            MsgBox("Error updating lab_id record." & vbNewLine & ex.Message.ToString)
                            Return 0
                        End Try
                    Case Else
                        MsgBox("Error updating lab_id record." & vbNewLine & "Multiple lab_id records found.")
                        Return 0
                End Select
            Else
                MsgBox("Error updating lab_id record." & vbNewLine & "Problem with lab_id table.")
                Return 0
            End If
        Catch ex As Exception
            MsgBox("Error updating lab_id record" & vbNewLine & ex.Message.ToString)
            Return 0
        End Try

    End Function

End Class
