Imports System.IO

Public Class Form
    Const endOfHeaders As String = vbCrLf

    Public Shared Function Main(ByVal CmdArgs() As String) As Integer
        Dim contentLength As Long = CLng(Environment.GetEnvironmentVariable("CONTENT_LENGTH"))

        Dim contentType As String = Environment.GetEnvironmentVariable("CONTENT_TYPE")
        ' 415 Unsupported Media Type

        Dim inputFileNameExtension As String = ".dwg"

        Dim tempDirectoryPath As String = Path.Combine(
            Path.GetTempPath(), Path.GetRandomFileName()
        )

        Dim tempDirectoryInfo As DirectoryInfo =
            Directory.CreateDirectory(tempDirectoryPath)

        Dim inputFileName =
            Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) _
            & inputFileNameExtension

        Dim inputFilePath As String = Path.Combine(
            tempDirectoryPath, inputFileName
        )

        Dim outputFileName =
            Path.GetFileNameWithoutExtension(Path.GetRandomFileName()) _
            & ".pdf"

        Dim outputFilePath As String = Path.Combine(
            tempDirectoryPath, outputFileName
        )

        Dim stdinStream As Stream = Console.OpenStandardInput()
        Dim stdoutStream As Stream = Console.OpenStandardOutput()
        Dim stderrStream As Stream = Console.OpenStandardError()

        Dim inputFileStream As FileStream = File.OpenWrite(inputFilePath)
        stdinStream.CopyTo(inputFileStream)
        inputFileStream.Close()

        Dim form As Form = New Form(outputFilePath)
        form.bravaComponent.Filename = inputFilePath

        Do While form.bravaComponent.Filename = inputFilePath
            System.Threading.Thread.Sleep(50)
        Loop

        If form.bravaComponent.ErrorMessage <> "No Error" Then
            tempDirectoryInfo.Delete(True)
            Return ReportError(form.bravaComponent.ErrorMessage,
                               stdoutStream, stderrStream)
        End If

        PrintUTF8ToStream(stdoutStream, "Content-Type: application/pdf" & vbCrLf)
        PrintUTF8ToStream(stdoutStream, endOfHeaders)

        Dim outputFileStream As FileStream = File.OpenRead(outputFilePath)
        outputFileStream.CopyTo(stdoutStream)
        outputFileStream.Close()

        tempDirectoryInfo.Delete(True)

        Return 0
    End Function

    Private Shared Function ReportError(
        ByVal message As String,
        ByRef stdoutStream As Stream,
        ByRef stderrStream As Stream
    ) As Integer
        Dim statusHeader As String = "Status: 501 Internal Server Error" & vbCrLf
        Dim contentTypeHeader As String = "Content-Type: text/plain" & vbCrLf

        PrintUTF8ToStream(stdoutStream, statusHeader)
        PrintUTF8ToStream(stdoutStream, contentTypeHeader)
        PrintUTF8ToStream(stdoutStream, endOfHeaders)
        PrintUTF8ToStream(stdoutStream, message)

        ' PrintUTF8ToStream(stderrStream, message)
        Return 0
    End Function

    Private Shared Sub PrintUTF8ToStream(ByRef stream As Stream, ByVal str As String)
        Dim strBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(str)
        stream.Write(strBytes, 0, strBytes.Length)
    End Sub


    Private errorMessage As String, outputFilePath As String

    Public Sub New(ByVal outputFilePath As String)
        InitializeComponent()

        Me.outputFilePath = outputFilePath
    End Sub

    Sub OnFileLoaded(ByVal sender As Object,
                     ByVal event_ As AxBRAVADTXLib._IBravaDTXViewEvents_FileLoadedEvent) _
    Handles bravaComponent.FileLoaded
        bravaComponent.ExportPDF(outputFilePath, 1)
    End Sub

    Sub OnFileLoadFailure(ByVal sender As Object,
                          ByVal event_ As AxBRAVADTXLib._IBravaDTXViewEvents_FileLoadFailureEvent) _
    Handles bravaComponent.FileLoadFailure
        errorMessage = "FileLoadFailure: " & event_.ToString()
    End Sub

    Sub OnExportPDFSuccess(ByVal sender As Object,
                           ByVal event_ As AxBRAVADTXLib._IBravaDTXViewEvents_ExportPDFSuccessEvent) _
    Handles bravaComponent.ExportPDFSuccess
        bravaComponent.CloseFile()
    End Sub

    Sub OnExportPDFFailure(ByVal sender As Object,
                           ByVal event_ As AxBRAVADTXLib._IBravaDTXViewEvents_ExportPDFFailureEvent) _
    Handles bravaComponent.ExportPDFFailure
        errorMessage = "ExportPDFFailure: " & event_.ToString()
    End Sub
End Class
