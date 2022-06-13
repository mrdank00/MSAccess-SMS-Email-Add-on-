Imports System.Data.OleDb
Imports System.Net
Imports CrystalDecisions.CrystalReports.Engine
Module SMSCRUD
    Public result As String
    Public cmd As New OleDbCommand
    Public da As New OleDbDataAdapter
    Public tbl As New DataTable
    Public ds As New DataSet
    Public dr As OleDbDataReader
    'Public dt As Fleetds


    Public Sub Insert(ByVal sql As String)
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            'HOLDS THE DATA TO BE EXECUTED
            With cmd
                .Connection = con
                .CommandText = sql
                'EXECUTE THE DATA
                result = cmd.ExecuteNonQuery
                'CHECKING IF THE DATA HAS EXECUTED OR NOT AND THEN THE POP UP MESSAGE WILL APPEAR
                If result = 0 Then
                    MessageBox.Show("Failed.", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    'MessageBox.Show("Success", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End With
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()
        End Try
    End Sub

    Public Sub Display(ByVal sql As String, dgv As DataGridView)
        'Try
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If
        'HOLDS THE DATA TO BE EXECUTED
        With cmd
            .Connection = con
            .CommandText = sql
            'EXECUTE THE DATA
            result = cmd.ExecuteNonQuery
            tbl = New DataTable
            da = New OleDbDataAdapter(cmd)
            da.Fill(tbl)
            dgv.DataSource = tbl
        End With
        con.Close()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'Finally

        'End Try
    End Sub
    Public Sub ComboFeed(ByVal sql As String, combo As ComboBox, row As Integer)
        If con.State = ConnectionState.Closed Then
            con.Open()
        End If
        With cmd
            .Connection = con
            .CommandText = sql
            dr = cmd.ExecuteReader
            combo.Items.Clear()
            While dr.Read
                combo.Items.Add(dr(row))
            End While

        End With
        con.Close()
    End Sub

    'Public Sub Report(ByVal sql As String, ByVal sql2 As String, ByVal tbl As String, ByVal tbl2 As String, rpt As ReportClass, rptviewer As CrystalReportViewer)
    '    Try

    '        con.Open()
    '        cmd = New SqlCommand(sql, con)
    '        dt.Tables(tbl).Rows.Clear()
    '        da.SelectCommand = cmd
    '        da.Fill(dt, tbl)

    '        dt.Tables(tbl2).Rows.Clear()
    '        cmd = New SqlCommand(sql2, con)
    '        da.SelectCommand = cmd
    '        da.Fill(dt, tbl2)

    '        'Dim report As New rptSalesPerDate
    '        rpt.SetDataSource(dt)
    '        rptviewer.ReportSource = rpt
    '        rptviewer.Refresh()
    '        cmd.Dispose()
    '        da.Dispose()
    '        con.Close()
    '    Catch ex As Exception
    '        MsgBox(ex.ToString)
    '    End Try
    'End Sub
    Public Sub checkbalance(ByVal smskey As String)
        Dim apikey = smskey
        ' Dim strGet As String
        Dim urlc As String = "https://sms.arkesel.com/sms/api?action=check-balance&api_key='" + apikey + "'&response=json"

        ' strGet = urlc
        Dim webclient As New System.Net.WebClient
        Dim result As String = webclient.DownloadString(urlc)
        MessageBox.Show(result)
    End Sub
    Public Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function
    Public Sub Sendsmsmessage(ByVal smskey As String, ByVal from As String, ByVal recipient As String, ByVal body As String)
        Dim apikey = smskey
        Dim message = body
        Dim numbers = recipient
        Dim strGet As String
        Dim sender = from
        Dim url As String = "https://sms.arkesel.com/sms/api?action=send-sms"
        strGet = url + "&api_key=" + apikey + "&to=" + numbers + "&from=" + sender + "&sms=" + WebUtility.UrlEncode(message)
        Dim webclient As New System.Net.WebClient
        Dim result As String = webclient.DownloadString(strGet)
        MessageBox.Show(result)
    End Sub
End Module
