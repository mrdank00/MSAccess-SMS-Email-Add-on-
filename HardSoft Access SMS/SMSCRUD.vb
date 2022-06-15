Imports System.Data.OleDb
Imports System.Net
Imports CrystalDecisions.CrystalReports.Engine
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Imports System.Net.WebClient
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

    Public Sub Display(ByVal sql As String, dgv As DataGridView, lbl As Label)
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
            lbl.Text = dgv.Rows.Count
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
    Public Sub checkbalance(ByVal smskey As String, lbl As Label)
        If CheckForInternetConnection() = True Then
            Dim apikey = smskey
            Dim urlc As String = "https://sms.arkesel.com/sms/api?action=check-balance&api_key=" + apikey + "&response=json"
            Dim webclient As New System.Net.WebClient
            Dim result As String = webclient.DownloadString(urlc)
            Dim mystring As String
            mystring = JsonConvert.SerializeObject(result, Formatting.Indented)

            lbl.Text = result

            'Dim mytext As String = result
            'Dim myChars() As Char = mytext.ToCharArray()
            'For Each ch As Char In myChars
            '    If Char.IsDigit(ch) Then
            '        Dim arr As Integer() = {}
            '        Dim newItem As String = ch

            '        arr = arr.Concat({newItem}).ToArray

            '        For index As Integer = 0 To arr.Length - 1
            '            MsgBox($"index: {index}, value: {arr(index)}")
            '        Next

            '    End If
            'Next
            'Dim mytext As String = result
            'Dim myChars() As Char = mytext.ToCharArray()
            'For Each ch As Char In myChars
            '    If Char.IsDigit(ch) Then

            '        lbl.Text = ch
            '    End If
            'Next
        End If


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
        'HttpWebRequest.Headers.Add("Header1", "Header1 value")
        ' strGet = "https://smsc.hubtel.com/v1/messages/send?clientsecret=zbuugivo&clientid=kfetlmjz&from=Hardsoft&to=233242838080&content=This+Is+A+Test+Message"
        Dim url As String = "https://sms.arkesel.com/sms/api?action=send-sms"
        strGet = url + "&api_key=" + apikey + "&to=" + numbers + "&from=" + sender + "&sms=" + WebUtility.UrlEncode(message)
        Dim webclient As New System.Net.WebClient
        Dim result As String = webclient.DownloadString(strGet)
        MessageBox.Show(result)
    End Sub
    Class Program
        Shared Function Main(ByVal args As String()) As Integer
            Dim array As Integer() = ExtractIntegers("animals|3|2|1|3")
            For Each i In array
                Console.WriteLine(i)
            Next
            Return 0
        End Function
        Shared Function ExtractIntegers(ByVal input As String) As Integer()
            Dim pattern As String = "animals(\|(?<number>[0-9]+))*"
            Dim match As Match = Regex.Match(input, pattern)
            Dim list As New List(Of Integer)
            If match.Success Then
                For Each capture As Capture In match.Groups("number").Captures
                    list.Add(Integer.Parse(capture.Value))
                Next
            End If
            Return list.ToArray()
        End Function
    End Class
End Module
