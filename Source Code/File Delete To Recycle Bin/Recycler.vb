Option Strict Off
Option Explicit On 
Imports System.IO
Imports System.Reflection
Module Recycler
    Private Structure SHFILEOPSTRUCT
        Dim hwnd As Integer
        Dim wFunc As Integer
        Dim pFrom As String
        Dim pTo As String
        Dim fFlags As Short
        Dim fAnyOperationsAborted As Boolean
        Dim hNameMappings As Integer
        Dim lpszProgressTitle As String
    End Structure

    Private Const FO_DELETE As Short = &H3S
    Private Const FOF_ALLOWUNDO As Short = &H40S
    Private Const FOF_NOCONFIRMATION As Short = &H10S

    Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
    "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Integer

    Public Function Recycle(ByRef sPath As String) As Integer
        Dim FileOp As SHFILEOPSTRUCT
        If Not File.Exists(sPath) Then
            MsgBox("Could not find " & sPath & "." & vbCrLf _
            & "Please verify the path.")
            Recycle = -1
            Exit Function
        End If
        With FileOp
            .wFunc = FO_DELETE
            .pFrom = sPath & vbNullChar
            .pTo = vbNullChar
            .fFlags = FOF_NOCONFIRMATION Or FOF_ALLOWUNDO
            .lpszProgressTitle = "Sending " & sPath & " to the Recycle Bin"
        End With
        Try
            SHFileOperation(FileOp)
        Catch ex As Exception
            Error_Handler(ex, "SHFileOperation")
        End Try
        Recycle = 0
    End Function

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
End Module

