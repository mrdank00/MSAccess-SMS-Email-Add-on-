Imports System.Data.OleDb
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
End Module
