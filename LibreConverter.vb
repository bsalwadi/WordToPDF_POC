Imports System.Diagnostics
Imports System.IO

Public Class LibreOfficeConverter
    Private ReadOnly _libreOfficePath As String
    Private ReadOnly _logFilePath As String

    Public Sub New(Optional libreOfficePath As String = "soffice", Optional logFilePath As String = "conversion.log")
        _libreOfficePath = libreOfficePath
        _logFilePath = logFilePath
    End Sub

    ''' <summary>
    ''' Converts a DOCX file to PDF using LibreOffice in headless mode.
    ''' </summary>
    Public Function ConvertDocxToPdf(inputFile As String, outputDir As String) As Boolean
        Try
            If Not File.Exists(inputFile) Then
                Log($"Input file not found: {inputFile}")
                Return False
            End If

            If Not Directory.Exists(outputDir) Then
                Directory.CreateDirectory(outputDir)
            End If

            Dim arguments As String = $"--headless --convert-to pdf --outdir ""{outputDir}"" ""{inputFile}"""
            Dim startInfo As New ProcessStartInfo With {
                .FileName = _libreOfficePath,
                .Arguments = arguments,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Log($"Starting conversion: {arguments}")

            Using process As Process = Process.Start(startInfo)
                Dim output As String = process.StandardOutput.ReadToEnd()
                Dim [error] As String = process.StandardError.ReadToEnd()
                process.WaitForExit()

                Log($"Standard Output: {output}")
                If Not String.IsNullOrEmpty([error]) Then
                    Log($"Standard Error: {[error]}")
                End If

                If process.ExitCode <> 0 Then
                    Log($"LibreOffice conversion failed with exit code {process.ExitCode}")
                    Return False
                End If

                Log("Conversion completed successfully.")
                Return True
            End Using

        Catch ex As Exception
            Log($"Exception occurred: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Logs message to a file with timestamp.
    ''' </summary>
    Private Sub Log(message As String)
        Dim logMessage As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}"
        File.AppendAllText(_logFilePath, logMessage & Environment.NewLine)
    End Sub
End Class

