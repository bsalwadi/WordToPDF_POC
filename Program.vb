Imports System

Module Program
    Sub Main(args As String())
        Dim converter As New LibreOfficeConverter("C:\Program Files\LibreOffice\program\soffice.exe", "C:\Users\Bassu\Downloads\conversion.log")

        Dim inputDoc As String = args(0)
        Dim outputDir As String = "C:\Users\Bassu\Downloads"

        If converter.ConvertDocxToPdf(inputDoc, outputDir) Then
            Console.WriteLine("Conversion successful.")
        Else
            Console.WriteLine("Conversion failed. Check log for details.")
        End If
    End Sub
End Module
