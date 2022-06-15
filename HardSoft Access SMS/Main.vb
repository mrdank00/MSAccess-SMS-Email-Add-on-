Imports System.Net.Mail
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net
Imports Newtonsoft.Json
Public Class Main
    Dim dt As New dsTerminal

    'Dim cryRpt As New ReportDocument
    'Dim pdfFile As String = "K:\Users\KISSI\Desktop\avareport.pdf"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Panel2.Hide()
            If txtGeneralMsg.Text = "" Then
                MsgBox("Kindly type in message")
                Exit Sub
            End If
            If CheckForInternetConnection() = True Then
                ProgressBar1.Maximum = DataGridView1.Rows.Count - 1
                For Each row As DataGridViewRow In DataGridView1.Rows
                    'Panel2.Visible = True
                    Panel2.Show()
                    Sendsmsmessage("U2xUcFVwVVdKZmxJQnhyTllWYWs", "PPS MADINA", row.Cells(3).Value, txtGeneralMsg.Text)

                    lblsent.Text = row.Index.ToString
                    lblrem.Text = DataGridView1.Rows.Count - 1 - row.Index.ToString
                    ProgressBar1.Value = Val(row.Index.ToString)
                Next
                txtGeneralMsg.Text = ""
                Panel2.Hide()
                'Panel2.Visible = False
                MsgBox("Message Sucessfully sent")
                checkbalance("U2xUcFVwVVdKZmxJQnhyTllWYWs", Label23)
                Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView1, Label25)
            Else
                MsgBox("NO INTERNET CONNECTION. KINDLY CONNECT TO THE INTERNET TO CONTINUE")
                Exit Sub
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub cbClassGeneral_MouseEnter(sender As Object, e As EventArgs) Handles cbClassGeneral.MouseEnter
        ComboFeed("SELECT * FROM tbSchClassID", cbClassGeneral, 2)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        onloads()
    End Sub
    Public Sub onloads()
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView1, Label25)
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE>0 and hometelno <> '" + "" + "' ", DataGridView2, Label26)
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE>0 and hometelno <> '" + "" + "' ", DataGridView4, Label26)
        Display("SELECT STUDID,STUDNAME,CLASS,totamtpaid,BalBFwd,BalCFwd,DATE,PayMode,ReceiptNo,HomeTelNo,Bankname,Narration FROM tbSchSmS_ReceiptPRINTED where hometelno<>'" + "" + "' ", DataGridView3, Label28)
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,arrears,currenttermbill,balance,currentterm  FROM usystbSchAdmissionDATA where hometelno<>'" + "" + "' ", DataGridView5, Label30)

        'Display("SELECT STUDID,STUDNAME,CLASS,Amount,BalBFwd,BalCFwd,DATE,PayMode FROM  usystbSchTerminalDataSHEET ", DataGridView3)
    End Sub

    Private Sub cbClassGeneral_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbClassGeneral.SelectedIndexChanged
        If cbClassGeneral.SelectedIndex <> -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where Presentclass='" + cbclassgeneral.Text + "' and hometelno <> '" + "" + "'", DataGridView1, Label25)
        End If
    End Sub

    Private Sub cbClassGeneral_KeyUp(sender As Object, e As KeyEventArgs) Handles cbClassGeneral.KeyUp
        If cbClassGeneral.SelectedIndex = -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView1, Label25)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Try
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            'email=39
            Dim query = "SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,EmailAddress FROM usystbSchAdmissionDATA "
            'where Emailaddress<>'" + "" + "' 
            cmd = New OleDb.OleDbCommand(query, con)
            da = New OleDb.OleDbDataAdapter(cmd)
            tbl = New DataTable()
            da.Fill(tbl)
            ' con.Close()

            For k = 0 To tbl.Rows.Count - 1

                cmd = New OleDb.OleDbCommand("select * from usystbSchTerminalDataSHEET where admissionno='" + tbl.Rows(k)(0).ToString + "' ", con)
                dt.Tables("usystbSchTerminalDataSHEET").Rows.Clear()
                da.SelectCommand = cmd
                da.Fill(dt, "usystbSchTerminalDataSHEET")


                'Dim sql = "select * from ClientReg"
                'dt.Tables("ClientReg").Rows.Clear()
                'cmd = New SqlCommand(sql, FleetCon)
                'da.SelectCommand = cmd
                'da.Fill(dt, "ClientReg")

                Dim report As New rptTerminalReport
                report.SetDataSource(dt)
                CrystalReportViewer1.ReportSource = report
                CrystalReportViewer1.Refresh()

                Dim pdfFile As String = "K:\Users\KISSI\Desktop\" + tbl.Rows(k)(1).ToString + " " + "Terminal Report" + ".pdf"

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

                'sendMail(tbl.Rows(k)(4).ToString, "TERMINAL REPORTS ", " Find below attached Terminal REPORT", pdfFile)
                'Sendsmsmessage("OlpRclJkOFBpUTZadVI5WWE=", "CTK School", tbl.Rows(k)(3).ToString, "The terminal Report for the academic Term has been sent to you By mail")
                report.Close()
                report.Dispose()
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            If ComboBox1.SelectedIndex = 0 Or ComboBox1.SelectedIndex = 1 Then
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                'email=39
                'Dim query = "SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,EmailAddress FROM usystbSchAdmissionDATA "
                Dim query = "SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,EmailAddress FROM usystbSchAdmissionDATA where course='" + ComboBox1.Text + "' "
                'where Emailaddress<>'" + "" + "' 
                cmd = New OleDb.OleDbCommand(query, con)
                da = New OleDb.OleDbDataAdapter(cmd)
                tbl = New DataTable()
                da.Fill(tbl)


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

                    Dim pdfFile As String = "K:\Users\KISSI\Desktop\New folder\" + ComboBox1.Text + "\" + tbl.Rows(k)(1).ToString + " Terminal Bill" + ".pdf"

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
                    'If CheckForInternetConnection() = True Then
                    '    sendMail(tbl.Rows(k)(4).ToString, "TERMINAL BILLS ", " Find below attached Terminal Bill", pdfFile)
                    '    Sendsmsmessage("OlpRclJkOFBpUTZadVI5WWE=", "CTK School", tbl.Rows(k)(3).ToString, "The terminal Bill for the academic year has been sent to you By mail")
                    'End If

                    report.Close()
                    report.Dispose()
                Next
            End If
            If ComboBox1.SelectedIndex = 2 Then
                If con.State = ConnectionState.Closed Then
                    con.Open()
                End If
                'email=39
                'Dim query = "SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,EmailAddress FROM usystbSchAdmissionDATA "
                Dim query = "SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,EmailAddress FROM usystbSchAdmissionDATA where course='" + ComboBox1.Text + "' "
                'where Emailaddress<>'" + "" + "' 
                cmd = New OleDb.OleDbCommand(query, con)
                da = New OleDb.OleDbDataAdapter(cmd)
                tbl = New DataTable()
                da.Fill(tbl)


                For k = 0 To tbl.Rows.Count - 1
                    Dim quer = "select * from tbSchBillingTRANX_SemesteR where studno='" + tbl.Rows(k)(0).ToString + "' "
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

                    Dim pdfFile As String = "K:\Users\KISSI\Desktop\New folder\" + ComboBox1.Text + "\" + tbl.Rows(k)(1).ToString + " Terminal Bill" + ".pdf"

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
                    'If CheckForInternetConnection() = True Then
                    '    sendMail(tbl.Rows(k)(4).ToString, "TERMINAL BILLS ", " Find below attached Terminal Bill", pdfFile)
                    '    Sendsmsmessage("OlpRclJkOFBpUTZadVI5WWE=", "CTK School", tbl.Rows(k)(3).ToString, "The terminal Bill for the academic year has been sent to you By mail")
                    'End If

                    report.Close()
                    report.Dispose()
                Next
            End If
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


    Private Sub cbclassgeneral_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles cbclassgeneral.SelectedIndexChanged

    End Sub

    Private Sub txtamt_TextChanged(sender As Object, e As EventArgs) Handles txtamt.TextChanged

        If txtamt.Text <> "" And cbsign.SelectedIndex <> -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE " + cbsign.Text + "" + txtamt.Text + " and hometelno <> '" + "" + "' ", DataGridView2, Label26)
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE " + cbsign.Text + "" + txtamt.Text + " and hometelno <> '" + "" + "' ", DataGridView4, Label26)
        Else
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView2, Label26)
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView4, Label26)
        End If
    End Sub

    Private Sub cbsign_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbsign.SelectedIndexChanged

        If txtamt.Text <> "" And cbsign.SelectedIndex <> -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE " + cbsign.Text + "" + txtamt.Text + " and hometelno <> '" + "" + "' ", DataGridView2, Label26)
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE " + cbsign.Text + "" + txtamt.Text + " and hometelno <> '" + "" + "' ", DataGridView4, Label26)
        Else
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView2, Label26)
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView4, Label26)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Panel2.Hide()
            If txtReminder.Text = "" Then
                MsgBox("Kindly type in message")
                Exit Sub
            End If
            If CheckForInternetConnection() = True Then
                Dim i As Integer = -1
                For Each row As DataGridViewRow In DataGridView2.Rows
                    'MsgBox(row.Index.ToString)
                    If i <> DataGridView1.Rows.Count Then
                        DataGridView4.Rows.RemoveAt(i + 1)
                    End If

                    Panel3.Show()
                    Sendsmsmessage("U2xUcFVwVVdKZmxJQnhyTllWYWs", "PPS MADINA", row.Cells(3).Value, "REMINDER" + vbNewLine + row.Cells(1).Value + vbNewLine + "Student ID:" + " " + row.Cells(0).Value & vbNewLine + row.Cells(2).Value + vbNewLine + "Bal. Due:" + " GHC " & row.Cells(4).Value & vbNewLine + "---------------" + vbNewLine + "NOTE: " + vbNewLine + txtReminder.Text)

                    Label5.Text = row.Index.ToString
                    Label4.Text = DataGridView2.Rows.Count - 1 - row.Index.ToString
                    ProgressBar2.Maximum = DataGridView2.Rows.Count - 1
                    ProgressBar2.Value = Val(row.Index.ToString)

                    ' DataGridView2.DataSource = DataGridView4
                Next
                txtReminder.Text = ""
                checkbalance("U2xUcFVwVVdKZmxJQnhyTllWYWs", Label23)
                Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE>0 and hometelno <> '" + "" + "' ", DataGridView2, Label26)
                Panel3.Hide()
                'Panel2.Visible = False
                MsgBox("Sucessfully sent")

            Else
                MsgBox("NO INTERNET CONNECTION. KINDLY CONNECT TO THE INTERNET TO CONTINUE")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs)
        Try
            'Dim pdate As Date = CDate(DateTimePicker1.Value)
            'MsgBox(DateTimePicker1.Value)
            'Display("SELECT STUDID,STUDNAME,CLASS,Amount,BalBFwd,BalCFwd,DATE,PayMode FROM usystbSchReceiptsPosted  WHERE date=12/11/2021  ", DataGridView3)
            '" & DateTimePicker1.Value & "
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Panel4.Hide()
            If DataGridView3.Rows.Count <> 0 Then
                If CheckForInternetConnection() = True Then
                    Dim i As Integer = -1
                    For Each row As DataGridViewRow In DataGridView3.Rows
                        'MsgBox(row.Index.ToString)
                        'If i <> DataGridView1.Rows.Count Then
                        '    DataGridView3.Rows.RemoveAt(i + 1)
                        'End If

                        Panel4.Show()
                        Sendsmsmessage("U2xUcFVwVVdKZmxJQnhyTllWYWs", "PPS MADINA", row.Cells(9).Value, "PAYMENT RECEIPT" + vbNewLine + "" + row.Cells(1).Value + vbNewLine + "Stud. ID:" + " " + row.Cells(0).Value & vbNewLine + "Class:" + " " + row.Cells(2).Value + vbNewLine + "--------------------" + vbNewLine + "Date:" + " " & row.Cells(6).Value & vbNewLine + "Receipt No:" + " " & row.Cells(8).Value & vbNewLine + "Bal. B/Fwd:" + " GHC " & row.Cells(4).Value & vbNewLine + "Amount Paid:" + " GHC " & row.Cells(3).Value & vbNewLine + "Bal. Payable:" + " GHC " & row.Cells(5).Value & vbNewLine + "----------------" + vbNewLine + "Bank Name: " + row.Cells(10).Value + vbNewLine + "Narration: " & row.Cells(11).Value)

                        Label10.Text = row.Index.ToString
                        Label9.Text = DataGridView3.Rows.Count - 1 - row.Index.ToString
                        ProgressBar3.Maximum = DataGridView3.Rows.Count - 1
                        ProgressBar3.Value = Val(row.Index.ToString)

                        Insert("delete from tbSchSmS_ReceiptPRINTED where receiptno='" + row.Cells(8).Value + "'")

                        'DataGridView2.DataSource = DataGridView4
                    Next
                    checkbalance("U2xUcFVwVVdKZmxJQnhyTllWYWs", Label23)
                    Display("SELECT STUDID,STUDNAME,CLASS,totamtpaid,BalBFwd,BalCFwd,DATE,PayMode,ReceiptNo,HomeTelNo,Bankname,Narration FROM tbSchSmS_ReceiptPRINTED where hometelno<>'" + "" + "' ", DataGridView3, Label28)
                    Display("SELECT STUDID,STUDNAME,CLASS,totamtpaid,BalBFwd,BalCFwd,DATE,PayMode,ReceiptNo,HomeTelNo,Bankname,Narration FROM tbSchSmS_ReceiptPRINTED where hometelno<>'" + "" + "' ", DataGridView3, Label28)
                    Panel4.Hide()
                    'Panel2.Visible = False
                    MsgBox("Sucess")

                Else
                    MsgBox("NO INTERNET CONNECTION. KINDLY CONNECT TO THE INTERNET TO CONTINUE")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Panel2.Hide()
        Panel3.Hide()
        Panel4.Hide()
        Panel5.Hide()

        'TabPage7.Hide()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO FROM usystbSchAdmissionDATA where hometelno <> '" + "" + "'", DataGridView1, Label25)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,BALANCE FROM usystbSchAdmissionDATA WHERE BALANCE>0 and hometelno <> '" + "" + "' ", DataGridView2, Label26)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Display("SELECT STUDID,STUDNAME,CLASS,totamtpaid,BalBFwd,BalCFwd,DATE,PayMode,ReceiptNo,HomeTelNo,Bankname,Narration FROM tbSchSmS_ReceiptPRINTED where hometelno<>'" + "" + "' ", DataGridView3, Label28)
    End Sub

    Private Sub TabControl1_Enter(sender As Object, e As EventArgs) Handles TabControl1.Enter
        onloads()
        checkbalance("U2xUcFVwVVdKZmxJQnhyTllWYWs", Label23)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            Panel5.Hide()
            If DataGridView5.Rows.Count <> 0 Then
                If CheckForInternetConnection() = True Then
                    Dim i As Integer = -1
                    For Each row As DataGridViewRow In DataGridView5.Rows
                        'MsgBox(row.Index.ToString)
                        'If i <> DataGridView1.Rows.Count Then
                        '    DataGridView3.Rows.RemoveAt(i + 1)
                        'End If

                        Panel5.Show()
                        Sendsmsmessage("U2xUcFVwVVdKZmxJQnhyTllWYWs", "PPS MADINA", row.Cells(3).Value, "TERMINAL BILL" + vbNewLine + "" + row.Cells(1).Value + vbNewLine + "Stud. ID:" + " " + row.Cells(0).Value & vbNewLine + "Class:" + " " + row.Cells(2).Value + vbNewLine + "--------------------" + vbNewLine + "Arrears:" + " Ghc " & row.Cells(4).Value & vbNewLine + "Current Term Bill:" + " Ghc " & row.Cells(5).Value & vbNewLine + "Balance:" + " Ghc " & row.Cells(6).Value & vbNewLine + "Bill Term : " + row.Cells(7).Value)

                        Label15.Text = row.Index.ToString
                        Label14.Text = DataGridView5.Rows.Count - 1 - row.Index.ToString
                        ProgressBar4.Maximum = DataGridView5.Rows.Count - 1
                        ProgressBar4.Value = Val(row.Index.ToString)

                    Next
                    checkbalance("U2xUcFVwVVdKZmxJQnhyTllWYWs", Label23)

                    Panel5.Hide()
                    'Panel2.Visible = False
                    MsgBox("Messages Sucessfully Sent")

                Else
                    MsgBox("NO INTERNET CONNECTION. KINDLY CONNECT TO THE INTERNET TO CONTINUE")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox2.MouseClick
        ComboFeed("SELECT * FROM tbSchClassID", ComboBox2, 2)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex <> -1 Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,arrears,currenttermbill,balance,currentterm  FROM usystbSchAdmissionDATA where hometelno<>'" + "" + "' and presentclass='" + ComboBox2.Text + "' ", DataGridView5, Label30)
        End If
    End Sub

    Private Sub ComboBox2_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox2.TextUpdate
        If ComboBox2.Text = "" Then
            Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,arrears,currenttermbill,balance,currentterm  FROM usystbSchAdmissionDATA where hometelno<>'" + "" + "' ", DataGridView5, Label30)
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Display("SELECT STUDID,STUDNAME,PRESENTCLASS,HOMETELNO,arrears,currenttermbill,balance,currentterm  FROM usystbSchAdmissionDATA where hometelno<>'" + "" + "' ", DataGridView5, Label30)
    End Sub
End Class
