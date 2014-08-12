Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public filequeue1 As String
    Private filecount As Long
    Private blankcount As Long
    Private fullcount As Long
    'Public searchterm As String
    Private filereader As StreamReader
    Private filewriter As StreamWriter

    Public Event WorkerFileProcessing(ByVal filename As String, ByVal queue As Integer)
    Public Event WorkerStatusMessage(ByVal message As String, ByVal statustag As Integer)
    Public Event WorkerError(ByVal Message As Exception)
    Public Event WorkerFileCount(ByVal Result As Long, ByVal StatusBar As Integer)
    Public Event WorkerComplete(ByVal queue As Integer)




#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As Exception)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message)
            End If
        Catch ex As Exception
            MsgBox("An error occurred in File Obfuscater's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal identifier_msg As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Error_Handler(exc)
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerFileObfuscate_Routine)
                    WorkerThread.Start()
                Case 2
                    filereader.Close()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub





    Private Sub WorkerFileObfuscate_Routine()
        RaiseEvent WorkerStatusMessage("Running File Obfuscation Operation", 1)
        Activity_Logger("-- Process Started --")
        filecount = 0
        fullcount = 0
        blankcount = 0

        RaiseEvent WorkerFileCount(filecount, 1)
        RaiseEvent WorkerFileCount(fullcount, 2)
        RaiseEvent WorkerFileCount(blankcount, 4)

        Try
            Dim finf As FileInfo = New FileInfo(filequeue1)
            'If MsgBox(finf.Length & " - " & System.Math.Round((finf.Length / 50)), MsgBoxStyle.OKCancel) = MsgBoxResult.OK Then

            Dim sizetouse As Long = CLng(System.Math.Round(finf.Length / 50))
            'MsgBox(sizetouse)
            If finf.Exists = True Then
                FileObfuscate(finf.FullName, sizetouse)
            End If
            'End If

            finf = Nothing
        Catch ex As Exception
            Error_Handler(ex)
        End Try

        ' RaiseEvent WorkerFileCount(filecount, 1)
        ' RaiseEvent WorkerFileCount(fullcount, 2)
        ' RaiseEvent WorkerFileCount(blankcount, 3)

        RaiseEvent WorkerStatusMessage("File Obfuscation Operation Completed", 1)
        Activity_Logger("-- Process Ended --")
        RaiseEvent WorkerComplete(0)
    End Sub

    Private Sub FileObfuscate(ByVal filename As String, ByVal buffersize_in As Long)
        Try
            Activity_Logger("File to Obfuscate: " & filename)
            RaiseEvent WorkerFileProcessing(filename, 1)
            Dim filecreated As ArrayList = New ArrayList

            Try
                RaiseEvent WorkerStatusMessage("Creating File Segments", 1)
                Dim finfo As FileInfo = New FileInfo(filename)
                Dim bytereader As StreamReader = New StreamReader(filename)
                Dim str As Stream = bytereader.BaseStream()
                'MsgBox(buffersize_in)
                Dim redimvalue As Long = buffersize_in
                'MsgBox(redimvalue)
                Dim inBuf(redimvalue) As Byte
                'MsgBox(inBuf.Length)
                Dim bytesToRead As Long
                Dim bytesRead As Long
                Dim totalbytesread As Long = 0
                Dim fstr As New FileStream(filename & "_.obf", FileMode.Create, FileAccess.Write)
                filecreated.Add(filename & "_.obf")

                Dim counter As Long = 0
                Dim fstrsplit As FileStream


                Try
                    While fstr.Length < finfo.Length
                        fstrsplit = New FileStream(filename & "" & counter & "_.obf", FileMode.Create, FileAccess.Write)
                        Activity_Logger(filename & "" & counter & "_.obf")
                        filecreated.Add(filename & "" & counter & "_.obf")
                        If finfo.Length - fstr.Length >= redimvalue Then
                            bytesToRead = redimvalue
                            ReDim inBuf(bytesToRead)
                        Else
                            bytesToRead = finfo.Length - fstr.Length
                            ReDim inBuf(bytesToRead)
                        End If
                        bytesRead = str.Read(inBuf, 0, bytesToRead)
                        totalbytesread = totalbytesread + bytesRead

                        fstr.Write(inBuf, 0, bytesRead)
                        fstrsplit.Write(inBuf, 0, bytesRead)
                        fstrsplit.Close()
                        counter = counter + 1
                        RaiseEvent WorkerFileCount(totalbytesread, 1)
                        RaiseEvent WorkerFileCount(counter, 2)
                    End While
                    RaiseEvent WorkerStatusMessage("Rebuilding File From Segments", 1)
                    Dim recounter As Long = 0
                    counter = counter - 1
                    Dim finfo2 As FileInfo
                    Dim bytereader2 As StreamReader
                    Dim str2 As Stream
                    For recounter = counter - 1 To 0 Step -1
                        finfo2 = New FileInfo(filename & "" & recounter & "_.obf")
                        bytereader2 = New StreamReader(finfo2.FullName)
                        str2 = bytereader2.BaseStream()
                        Dim inBuf2(redimvalue) As Byte
                        Dim bytesToRead2 As Long
                        Dim bytesRead2 As Long
                        Dim totalbytesread2 As Long = 0
                        Dim fstr2 As New FileStream(filename & ".obf", FileMode.Append, FileAccess.Write)
                        Dim counter2 As Long = 0
                        Dim displayunit2 As Integer
                        Dim displayunittext2 As String
                        bytesToRead2 = finfo2.Length
                        ReDim inBuf(bytesToRead2)
                        bytesRead2 = str2.Read(inBuf2, 0, bytesToRead2)
                        totalbytesread2 = totalbytesread2 + bytesRead2
                        fstr2.Write(inBuf2, 0, bytesRead2)
                        finfo2 = Nothing
                        str2.Close()
                        fstr2.Close()
                        RaiseEvent WorkerFileCount(recounter, 3)
                    Next

                    finfo2 = New FileInfo(filename & "" & counter & "_.obf")
                    bytereader2 = New StreamReader(finfo2.FullName)
                    str2 = bytereader2.BaseStream()
                    Dim inBuf3(redimvalue) As Byte
                    Dim bytesToRead3 As Long
                    Dim bytesRead3 As Long
                    Dim totalbytesread3 As Long = 0
                    Dim fstr3 As New FileStream(filename & ".obf", FileMode.Append, FileAccess.Write)
                    Dim counter3 As Long = 0

                    Dim displayunit3 As Integer
                    Dim displayunittext3 As String

                    bytesToRead3 = finfo2.Length
                    ReDim inBuf(bytesToRead3)

                    bytesRead3 = str2.Read(inBuf3, 0, bytesToRead3)
                    totalbytesread3 = totalbytesread3 + bytesRead3
                    fstr3.Write(inBuf3, 0, bytesRead3)
                    finfo2 = Nothing
                    str2.Close()
                    fstr3.Close()

                    If fstr.Length <> finfo.Length Then
                        'result = "Downloaded (Corrupt Download)"
                    End If
                Catch ex As Exception
                    'result = "Download Failed"
                    Error_Handler(ex)
                End Try

                finfo = Nothing
                str.Close()
                fstr.Close()

                RaiseEvent WorkerStatusMessage("Removing Temporary Files", 1)

                Dim tinf As FileInfo
                Dim runner As Integer = 1
                For Each stri As String In filecreated
                    tinf = New FileInfo(stri)
                    If tinf.Exists = True Then
                        tinf.Delete()
                    End If
                    RaiseEvent WorkerFileCount(runner, 4)
                    runner = runner + 1
                Next


            Catch ex As Exception
                Error_Handler(ex)
            End Try

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub 'ProcessDirectory

End Class
