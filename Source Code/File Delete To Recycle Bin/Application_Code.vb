Imports Microsoft.Win32
Imports System.IO
Imports System.Reflection

Module Application_Code

    Sub Main(ByVal sArgs() As String)
        Try
            If sArgs.Length < 1 Or sArgs.Length > 2 Then
                Console.WriteLine("File Delete To Recycle Bin")
                Console.WriteLine("-------------------------")
                Console.WriteLine("Usage: executable [FileToDelete] {Reporting}")
                Console.WriteLine("  where:")
                Console.WriteLine("   - [FileToDelete] is the full path of the file to delete")
                Console.WriteLine("   - {Reporting} Optional. 1 for full reporting, 0 for silent. Default is 0")
            Else
                Dim silent As Boolean = True
                If sArgs.Length = 2 Then
                    If sArgs(1) = "1" Then
                        silent = False
                    End If
                End If
                If silent = False Then
                    Console.WriteLine(DeleteFile(sArgs(0), silent))
                Else
                    DeleteFile(sArgs(0), silent)
                End If
            End If

        Catch ex As Exception
            Console.WriteLine("Fail. Check Error Log for more details.")
            Error_Handler(ex, "Main Code")
        End Try
    End Sub

    Private Function ApplicationPath() As String
        Return _
        Path.GetDirectoryName([Assembly].GetEntryAssembly().Location)
    End Function

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Console.WriteLine("An error occurred in File Delete To Recycle Bin's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Error_Handler(exc, "Activity_Logger")
        End Try
    End Sub

    Public Function DeleteFile(ByVal FileToDelete As String, ByVal silent As Boolean) As String
        Dim result As String = "Fail."
        Try
          
            Dim finfo As FileInfo = New FileInfo(FileToDelete)
            Dim recyclestatus As Integer = -1
            If finfo.Exists = True Then
                If silent = False Then
                    Activity_Logger("File to Delete: " & finfo.FullName)
                End If

                recyclestatus = Recycle(finfo.FullName)
                If recyclestatus = 0 Then
                    result = "File has been deleted"
                Else
                    result = "File deletion has failed"
                End If
            Else
                Activity_Logger("File to Delete: " & finfo.FullName)
                result = "File to delete cannot be located at '" & FileToDelete & "'"
            End If




            finfo = Nothing

        Catch ex As Exception
            Error_Handler(ex, "DeleteFile")
            result = "Fail. Check Error Log for further details"
        End Try
        Try
            If silent = False Then
                Activity_Logger("Result: " & result)
            End If
        Catch ex As Exception
            Error_Handler(ex, "DeleteFile")
            result = "Fail. Check Error Log for further details"
        End Try
        Return result
    End Function

End Module
