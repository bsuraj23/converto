Imports System.IO

Public Class Form
    Public Shared Function Main(ByVal CmdArgs() As String) As Integer
        Dim inputFileExtension As String = CmdArgs(0)
        Dim inputContentType As String = CmdArgs(1)
        Dim inputFilePath As String = CmdArgs(2)
        Dim outputFilePath As String = CmdArgs(3)

        Dim form As Form = New Form(outputFilePath)
        form.bravaComponent.Filename = inputFilePath

        Do While form.bravaComponent.Filename = inputFilePath
            System.Threading.Thread.Sleep(50)
        Loop

        If form.bravaComponent.ErrorMessage <> "No Error" Then
            Return ReportError(form.bravaComponent.ErrorMessage)
        End If

        Return 0
    End Function

    Private Shared Function ReportError(
        ByVal message As String
    ) As Integer
        Dim stderrStream As Stream = Console.OpenStandardError()

        PrintUTF8ToStream(stderrStream, message)

        Return 1
    End Function

    Private Shared Sub PrintUTF8ToStream(ByRef stream As Stream, ByVal str As String)
        Dim strBytes As Byte() = System.Text.Encoding.UTF8.GetBytes(str)
        stream.Write(strBytes, 0, strBytes.Length)
    End Sub


    Private outputFilePath As String ', errorMessage As String

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
        ' errorMessage = "FileLoadFailure: " & event_.ToString()
    End Sub

    Sub OnExportPDFSuccess(ByVal sender As Object,
                           ByVal event_ As AxBRAVADTXLib._IBravaDTXViewEvents_ExportPDFSuccessEvent) _
    Handles bravaComponent.ExportPDFSuccess
        bravaComponent.CloseFile()
    End Sub

    Sub OnExportPDFFailure(ByVal sender As Object,
                           ByVal event_ As AxBRAVADTXLib._IBravaDTXViewEvents_ExportPDFFailureEvent) _
    Handles bravaComponent.ExportPDFFailure
        ' errorMessage = "ExportPDFFailure: " & event_.ToString()
    End Sub
End Class
