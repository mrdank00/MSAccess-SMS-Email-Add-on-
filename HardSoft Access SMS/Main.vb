Imports System.Net.Mail
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net
Public Class Main
    Dim dt As New dsTerminal

    'Dim cryRpt As New ReportDocument
    'Dim pdfFile As String = "K:\Users\KISSI\Desktop\avareport.pdf"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each row As DataGridViewRow In DataGridView1.Rows
            Sendsmsmessage("OlpRclJkOFBpUTZadVI5WWE=", "CTK School", row.Cells(3).Value, txtGeneralMsg.Text)
        Next
        MsgBox("Sucess")
    End Sub

    Private Sub cbClassGeneral_MouseEnter(sender As Object, e As EventArgs) Handles cbClassGeneral.MouseEnter
        ComboFeed("SELECT * FROM tbSchClassID", cbClassGeneral, 2)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView1)
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE>0 and hometelno <> '" + "" + "' ", DataGridView2)
        Display("SELECT STUDID,STUDNAME,CLASS,Amount,BalBFwd,BalCFwd,DATE,PayMode FROM usystbSchReceiptsPosted  ", DataGridView3)
        'Display("SELECT STUDID,STUDNAME,CLASS,Amount,BalBFwd,BalCFwd,DATE,PayMode FROM  usystbSchTerminalDataSHEET ", DataGridView3)
    End Sub

    Private Sub cbClassGeneral_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbClassGeneral.SelectedIndexChanged
        If cbClassGeneral.SelectedIndex <> -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where Presentclass='" + cbclassgeneral.Text + "' and hometelno <> '" + "" + "'", DataGridView1)
        End If
    End Sub

    Private Sub cbClassGeneral_KeyUp(sender As Object, e As KeyEventArgs) Handles cbClassGeneral.KeyUp
        If cbClassGeneral.SelectedIndex = -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "' ", DataGridView1)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim query = "select * from usystbSchTERMINALREPORTdata "
        cmd = New OleDb.OleDbCommand(query, con)
        dt.Tables("TerminalReport").Rows.Clear()
        da.SelectCommand = cmd
        da.Fill(dt, "TerminalReport")


        'Dim sql = "select * from ClientReg"
        'dt.Tables("ClientReg").Rows.Clear()
        'cmd = New SqlCommand(sql, FleetCon)
        'da.SelectCommand = cmd
        'da.Fill(dt, "ClientReg")

        Dim report As New rptTerminalReport
        report.SetDataSource(dt)
        CrystalReportViewer1.ReportSource = report
        CrystalReportViewer1.Refresh()

        'Dim CrExportOptions As ExportOptions
        'Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
        'Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions
        'CrDiskFileDestinationOptions.DiskFileName = pdfFile
        'CrExportOptions = report.ExportOptions
        'With CrExportOptions
        '    .ExportDestinationType = ExportDestinationType.DiskFile
        '    .ExportFormatType = ExportFormatType.PortableDocFormat
        '    .DestinationOptions = CrDiskFileDestinationOptions
        '    .FormatOptions = CrFormatTypeOptions
        'End With
        'report.Export()
        'sendMail()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            'email=39
            Dim query = "SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,EmailAddress FROM usystbSchAdmissionDATA where Emailaddress<>'" + "" + "' "
            cmd = New OleDb.OleDbCommand(query, con)
            da = New OleDb.OleDbDataAdapter(cmd)
            tbl = New DataTable()
            da.Fill(tbl)
            con.Close()

            For k = 0 To tbl.Rows.Count - 1
                Dim quer = "select * from tbSchBillingTRANX where studno='" + tbl.Rows(k)(0).ToString + "' "
                cmd = New OleDb.OleDbCommand(quer, con)
                dt.Tables("TerminalBill").Rows.Clear()
                da.SelectCommand = cmd
                da.Fill(dt, "TerminalBill")


                'Dim sql = "select * from ClientReg"
                'dt.Tables("ClientReg").Rows.Clear()
                'cmd = New SqlCommand(sql, FleetCon)
                'da.SelectCommand = cmd
                'da.Fill(dt, "ClientReg")

                Dim report As New rptTerminalBill
                report.SetDataSource(dt)
                CrystalReportViewer2.ReportSource = report
                CrystalReportViewer2.Refresh()

                Dim pdfFile As String = "K:\Users\KISSI\Desktop\" + tbl.Rows(k)(1).ToString + "Terminal Bill" + ".pdf"

                Dim CrExportOptions As ExportOptions
                Dim CrDiskFileDestinationOptions As New DiskFileDestinationOptions()
                Dim CrFormatTypeOptions As New PdfRtfWordFormatOptions
                CrDiskFileDestinationOptions.DiskFileName = pdfFile
                CrExportOptions = report.ExportOptions
                With CrExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                    .DestinationOptions = CrDiskFileDestinationOptions
                    .FormatOptions = CrFormatTypeOptions
                End With
                report.Export()

                sendMail(tbl.Rows(k)(4).ToString, "TERMINAL BILLS ", " Find below attached Terminal Bill", pdfFile)
                Sendsmsmessage("OlpRclJkOFBpUTZadVI5WWE=", "CTK School", tbl.Rows(k)(3).ToString, "The terminal Bill for the academic year has been sent to you By mail")
                report.Close()
                report.Dispose()
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub sendMail(ByVal ToAddress As String, ByVal Subject As String, ByVal Body As String, ByVal pdffile As String)
        Try
            Dim mail As New MailMessage
            Dim smtpserver As New SmtpClient("smtp.gmail.com")
            mail.From = New MailAddress("kissidaniel7@gmail.com")
            mail.To.Add(ToAddress)
            mail.Subject = Subject
            mail.Body = Body
            mail.Attachments.Add(New Attachment(pdffile))

            smtpserver.Port = 587
            smtpserver.Credentials = New System.Net.NetworkCredential("kissidaniel7@gmail.com", "lienad00")
            smtpserver.EnableSsl = True
            smtpserver.Send(mail)
            MsgBox("Sent")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
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

    Private Sub cbclassgeneral_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cbclassgeneral.SelectedIndexChanged

    End Sub

    Private Sub txtamt_TextChanged(sender As Object, e As EventArgs) Handles txtamt.TextChanged
        If cbsign.SelectedIndex = -1 Then
            MsgBox("Kindly select Sign operator")
            cbsign.Focus()
        End If
        If txtamt.Text <> "" Or cbsign.SelectedIndex <> -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE " + cbsign.Text + "" + txtamt.Text + " and hometelno <> '" + "" + "' ", DataGridView2)
        End If
    End Sub
End Class
