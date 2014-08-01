' AddToRMA02032 - reads a list of serial numbers from Todd's file and adds to RMA02032
'
Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Console

Module Module1

    Dim ConnectionString As String

    Sub Main()
        Dim garbage As String
        ConnectionString = "DSN=PTS4"
        Console.WriteLine("Starting to update RMA02032 on " & Today.UtcNow.ToString & "(UTC)")
        Dim conn As New OdbcConnection
        conn.ConnectionString = ConnectionString
        Try
            conn.Open()
            UpdateRMA(conn)
        Catch myerror As OdbcException
            Console.WriteLine("Error connecting to the database: " & myerror.Message)
        Catch e As Exception
            Console.WriteLine("ERROR: " & e.Message)
        Finally
            If conn.State <> ConnectionState.Closed Then conn.Close()
        End Try
        Console.WriteLine("Completed updating to RMA02032 on " & Today.UtcNow.ToString & "(UTC)")
        Console.WriteLine("All done! Hit <Enter>")
        garbage = Console.ReadLine

    End Sub

    Private Sub UpdateRMA(ByVal myconn As OdbcConnection)
        Dim myCommand1 As New OdbcCommand
        Dim myCommand2 As New OdbcCommand
        Dim SQLstring As String
        Dim SerialNumberFile, SN As String
        Dim counter As Integer = 0

        Console.WriteLine("Enter file name to read: ")
        SerialNumberFile = Console.ReadLine()
        If System.IO.File.Exists(SerialNumberFile) = True Then
            Dim SNReader As New System.IO.StreamReader(SerialNumberFile)
            Do While SNReader.Peek() <> -1
                SN = SNReader.ReadLine()

                Try
                    ' First get all the test detail records for a given date
                    myCommand1.Connection = myconn
                    SQLstring = "Insert into repair_head values ('RMA02032','" & SN & "','2013-02-25 00:00:00','to be program ','','L','C','NONE','NONE',1,'2013-02-11 00:00:00',0,'0000-00-00 00:00:00','NO',14,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,NULL,1,0,NULL,0,0,0,NULL,0,0,'',0)"
                    myCommand1.CommandText = SQLstring
                    myCommand1.ExecuteNonQuery()
                    counter = counter + 1
                Catch e As Exception
                    Console.WriteLine("ERROR: " & e.Message)

                End Try
            Loop
            myconn.Close()
            SNReader.Close()
            Console.WriteLine(counter.ToString & " records were updated.")
        Else
            Console.WriteLine("Input file can not be opened")
        End If
    End Sub
End Module
