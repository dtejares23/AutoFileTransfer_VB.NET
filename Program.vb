Imports System
Imports System.IO
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Text.Json ' For reading JSON config

Module Program
    ' Importing required Windows API functions
    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function GetConsoleWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function GetSystemMenu(hWnd As IntPtr, bRevert As Boolean) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Function RemoveMenu(hMenu As IntPtr, nPosition As UInteger, wFlags As UInteger) As Boolean
    End Function

    ' Constants for system menu options
    Private Const SC_CLOSE As UInteger = &HF060UI
    Private Const MF_BYCOMMAND As UInteger = &H0UI

    ' Variables for paths (read from config)
    Private sourcePath As String
    Private destinationPath As String
    Private logPath As String

    Sub Main()
        ' Read configuration file
        If Not LoadConfig() Then
            'DisableCloseButton()
            CountdownBeforeExit(5)
        End If

        DisableCloseButton()
        Console.WriteLine("Close button disabled to avoid interruption." & $"{vbCrLf}*AUTOMATED FILE TRANSFER*")
        Log("Program started.")

        Try
            If Not Directory.Exists(sourcePath) Then
                Dim message = "Source path not found or can't connect to server. PLEASE CONTACT IT SUPPORT"
                Console.WriteLine(message)
                Log("ERROR: " & message)
                CountdownBeforeExit(5)
            End If

            If Not Directory.Exists(destinationPath) Then
                Try
                    Console.WriteLine("Destination path not found." & $"{vbCrLf}Creating destination folder...")
                    Directory.CreateDirectory(destinationPath)
                    If Not Directory.Exists(destinationPath) Then Throw New IOException("Failed to create destination folder.")
                    Log($"INFO: Destination folder created: {destinationPath}")
                Catch ex As Exception
                    Dim errorMessage = "ERROR: " & ex.Message
                    Console.WriteLine(errorMessage)
                    Log(errorMessage)
                    CountdownBeforeExit(5)
                End Try
            End If

            Dim files As String() = Directory.GetFiles(sourcePath)
            Dim fileCount As Integer = files.Length
            Dim copiedFile As Integer = 0
            Dim today = DateTime.Now.Date

            For Each filePath As String In files
                Dim fileCreationDate = File.GetCreationTime(filePath).Date
                If fileCreationDate = today Then
                    Dim fileName As String = Path.GetFileName(filePath)
                    Dim destinationFile As String = Path.Combine(destinationPath, fileName)
                    System.IO.File.Copy(filePath, destinationFile, True)
                    copiedFile += 1
                    Dim message = $"Copied file {copiedFile}/{fileCount}: {fileName}"
                    Console.WriteLine(message)
                    Log("INFO: " & message)
                End If
            Next



            Dim successMessage = $"{vbCrLf}File transfer complete. Total files copied: {copiedFile}{vbCrLf}"
            Console.WriteLine(successMessage)
            Log("INFO: " & successMessage)
            CountdownBeforeExit(3)
        Catch ex As Exception
            Dim errorMessage = "ERROR: An unexpected error occurred: " & ex.Message
            Console.WriteLine(errorMessage)
            Log(errorMessage)
            CountdownBeforeExit(5)
        End Try
    End Sub

    Sub DisableCloseButton()
        Dim consoleHandle As IntPtr = GetConsoleWindow()
        If consoleHandle <> IntPtr.Zero Then
            Dim sysMenu As IntPtr = GetSystemMenu(consoleHandle, False)
            If sysMenu <> IntPtr.Zero Then
                RemoveMenu(sysMenu, SC_CLOSE, MF_BYCOMMAND)
            End If
        End If
    End Sub

    Sub CountdownBeforeExit(seconds As Integer)
        For i As Integer = seconds To 1 Step -1
            Console.WriteLine($"Exiting in {i} seconds...")
            'Console.WriteLine("Bye!")
            Thread.Sleep(1000) ' Pause for 1 second
        Next
        Log($"Program exited.{vbCrLf}")
        Environment.Exit(0)
    End Sub

    Sub Log(message As String)
        Try
            If Not Directory.Exists(logPath) Then
                Directory.CreateDirectory(logPath)
            End If
            Dim logFile As String = Path.Combine(logPath, "Log_" & DateTime.Now.ToString("yyyy-MM-dd") & ".txt")
            Using writer As New StreamWriter(logFile, True)
                writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}")
            End Using
        Catch ex As Exception
            Console.WriteLine("Failed to write to log: " & ex.Message)
        End Try
    End Sub

    Function LoadConfig() As Boolean
        Dim configFilePath As String = "config.json"
        Try
            If Not File.Exists(configFilePath) Then Throw New FileNotFoundException("Configuration file not found.")
            Dim json As String = File.ReadAllText(configFilePath)
            Dim config = JsonSerializer.Deserialize(Of Dictionary(Of String, String))(json)

            sourcePath = config("sourcePath")
            destinationPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), config("destinationPath"))
            logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), config("logPath"))
            Return True
        Catch ex As Exception
            Console.WriteLine("Failed to load configuration: " & ex.Message)
            Return False
        End Try
    End Function
End Module
