﻿Imports System.Data.OleDb
Module SMScon
    Public path = "K:\Users\KISSI\Desktop\CLASS 2A.mdb"

    Public Function SMSConnection() As OleDbConnection

        Return New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + path + "' ;")
    End Function
    Public con As OleDbConnection = SMSConnection()
End Module
