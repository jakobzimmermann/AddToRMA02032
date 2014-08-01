Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Console

Module CS1120Failure
    ' CS1120Failure - reads a list of serial numbers from Todd's file and extracts test data
    '
    Dim ConnectionString As String

    Sub Main()
        Dim garbage As String
        ConnectionString = "DSN=PTS4"
        Console.WriteLine("Starting to query CS1120 List on " & Today.UtcNow.ToString & "(UTC)")
        Dim conn As New OdbcConnection
        conn.ConnectionString = ConnectionString
        Try
            conn.Open()
            FindCS1120(conn)
        Catch myerror As OdbcException
            Console.WriteLine("Error connecting to the database: " & myerror.Message)
        Catch e As Exception
            Console.WriteLine("ERROR: " & e.Message)
        Finally
            If conn.State <> ConnectionState.Closed Then conn.Close()
        End Try
        Console.WriteLine("Completed querying for CS1120 Failures on " & Today.UtcNow.ToString & "(UTC)")
        Console.WriteLine("All done! Hit <Enter>")
        garbage = Console.ReadLine

    End Sub

    Private Sub FindCS1120(ByVal myconn As OdbcConnection)
        Dim myCommand1 As New OdbcCommand
        Dim myCommand2 As New OdbcCommand
        Dim MyReader As OdbcDataReader
        Dim SQLstring As String
        Dim SerialNumberFile, SN As String
        Dim counter As Integer = 0

        Console.WriteLine("Enter file name to read: ")
        SerialNumberFile = Console.ReadLine()
        If System.IO.File.Exists(SerialNumberFile) = True Then
            Dim SNReader As New System.IO.StreamReader(SerialNumberFile)
            Dim OutputFile As New System.IO.StreamWriter("C:\temp\CS1120FailureDetail.csv")
            Dim temp As String
            Do While SNReader.Peek() <> -1
                SN = SNReader.ReadLine()

                Try
                    ' First get all the test detail records for a given date
                    myCommand1.Connection = myconn
                    SQLstring = "SELECT n.* FROM   prod_test_detail n INNER JOIN (SELECT MAX(test_date_time) AS MaxDate, Serial_number " & _
                        "FROM prod_test_detail GROUP BY serial_number) nm ON nm.maxdate = n.test_date_time " & _
                        "AND nm.serial_number = n.serial_number and n.serial_number='" & SN & "'"

                    myCommand1.CommandText = SQLstring
                    MyReader = myCommand1.ExecuteReader
                    MyReader.Read()
                    Console.Write(SN & "|" & MyReader.GetString(0) & "|")
                    OutputFile.Write(SN & "|" & MyReader.GetString(0) & "|")
                    For t = 2 To 22
                        temp = MyReader.GetString(t)
                        If temp.Length > 0 Then
                            For u = 1 To temp.Length
                                If Mid(temp, u, 1) = Chr(10) Or Mid(temp, u, 1) = Chr(13) Then
                                    Mid$(temp, u, 1) = " "
                                End If
                            Next
                        End If
                        Console.Write(temp & "|")
                        OutputFile.Write(temp & "|")
                    Next
                    Console.WriteLine("")
                    OutputFile.WriteLine("")
                    MyReader.Close()
                    counter = counter + 1
                Catch e As Exception
                    Console.WriteLine("ERROR: " & e.Message)

                End Try
            Loop
            myconn.Close()
            SNReader.Close()
            Console.WriteLine(counter.ToString & " records were queried.")
        Else
            Console.WriteLine("Input file can not be opened")
        End If
    End Sub
End Module

