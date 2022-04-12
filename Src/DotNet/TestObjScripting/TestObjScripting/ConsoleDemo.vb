Imports ObjScriptingLib
Imports PigCmdLib

Public Class ConsoleDemo
    Private oFS As New FileSystemObject
    Public PigConsole As New PigConsole
    Public FilePath As String
    Public FolderPath As String

    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to FileSystemObject")
            Console.WriteLine("Press B to TextStream")
            Console.WriteLine("Press C to Encoder")
            'Console.WriteLine("Press D to Dictionary")
            'Console.WriteLine("Press E to Drive")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Menu FileSystemObject")
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to GetFile")
                        Console.WriteLine("Press B to GetFolder")
                        Console.WriteLine("Press C to FileExists")
                        Console.WriteLine("Press D to FolderExists")
                        Console.WriteLine("Press E to CreateFolder")
                        Console.WriteLine("Press F to GetTempName")
                        Console.WriteLine("Press G to GetFileVersion")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey(True).Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.WriteLine("#################")
                                Me.PigConsole.GetLine("Enter the file path", Me.FilePath)
                                Dim oFile As File = oFS.GetFile(Me.FilePath)
                                Console.WriteLine("DateCreated: " & oFile.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFile.DateLastModified)
                                Console.WriteLine("Name: " & oFile.Name)
                                Console.WriteLine("Path: " & oFile.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Console.WriteLine("#################")
                                Me.PigConsole.GetLine("Enter the file folder", Me.FolderPath)
                                Dim oFolder As Folder = oFS.GetFolder(Me.FolderPath)
                                Console.WriteLine("DateCreated: " & oFolder.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFolder.DateLastModified)
                                Console.WriteLine("Name: " & oFolder.Name)
                                Console.WriteLine("Path: " & oFolder.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.C
                                Me.PigConsole.GetLine("Enter the file path", Me.FilePath)
                                Console.WriteLine("FileExists: " & oFS.FileExists(Me.FilePath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.D
                                Me.PigConsole.GetLine("Enter the file folder", Me.FolderPath)
                                Console.WriteLine("FolderExists: " & oFS.FolderExists(Me.FolderPath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.E
                                Me.PigConsole.GetLine("Enter the file folder", Me.FolderPath)
                                oFS.CreateFolder(Me.FolderPath)
                                Console.Write("CreateFolder: ")
                                If oFS.LastErr = "" Then
                                    Console.WriteLine("OK")
                                Else
                                    Console.WriteLine(oFS.LastErr)
                                End If
                                Console.WriteLine("#################")
                            Case ConsoleKey.F
                                Console.WriteLine("#################")
                                Console.WriteLine("GetTempName: " & oFS.GetTempName)
                                Console.WriteLine("#################")
                            Case ConsoleKey.G
                                Console.WriteLine("#################")
                                Me.PigConsole.GetLine("Enter the file path", Me.FilePath)
                                Console.WriteLine("GetFileVersion: " & oFS.GetFileVersion(Me.FilePath))
                                Console.WriteLine("#################")
                        End Select
                    Loop
                Case ConsoleKey.B
                    Do While True
                        Console.WriteLine("*******************")
                        Console.WriteLine("Menu TextStream")
                        Console.WriteLine("*******************")
                        Console.WriteLine("Press Q to Up")
                        Console.WriteLine("Press A to Read File")
                        Console.WriteLine("Press B to Write File")
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey().Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.WriteLine("#################")
                                Me.PigConsole.GetLine("Enter the file path", Me.FilePath)
                                If oFS.FileExists(Me.FilePath) = False Then
                                    Console.WriteLine(Me.FilePath & " not found.")
                                Else
                                    Dim oTextStream As TextStream
                                    Console.WriteLine("OpenTextFile(" & Me.FilePath & ")...")
                                    oTextStream = oFS.OpenTextFile(Me.FilePath, FileSystemObject.IOMode.ForReading, False)
                                    If oFS.LastErr <> "" Then
                                        Console.WriteLine(oFS.LastErr)
                                    Else
                                        Do While Not oTextStream.AtEndOfStream
                                            Console.WriteLine(oTextStream.ReadLine)
                                        Loop
                                        oTextStream.Close()
                                    End If
                                End If
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Me.PigConsole.GetLine("Enter the file path", Me.FilePath)
                                Dim oTextStream As TextStream
                                Console.WriteLine("OpenTextFile(" & Me.FilePath & ")...")
                                oTextStream = oFS.OpenTextFile(Me.FilePath, FileSystemObject.IOMode.ForWriting, True)
                                If oFS.LastErr <> "" Then
                                    Console.WriteLine(oFS.LastErr)
                                Else
                                    oTextStream.WriteLine("WriteLine")
                                    oTextStream.WriteBlankLines(2)
                                    oTextStream.Close()
                                    Console.WriteLine("OK")
                                End If
                        End Select
                    Loop
                Case ConsoleKey.C
                    Console.WriteLine("*******************")
                    Console.WriteLine("Menu Encoder")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input VBS file path:")
                    Dim strVBSFilePath As String = Console.ReadLine
                    If oFS.FileExists(strVBSFilePath) = False Then
                        Console.WriteLine(strVBSFilePath & " not found.")
                    Else
                        Dim oEncoder As New Encoder
                        Console.WriteLine("OpenTextFile(" & strVBSFilePath & ")")
                        Dim tsIn As TextStream = oFS.OpenTextFile(strVBSFilePath, FileSystemObject.IOMode.ForReading, False)
                        If oFS.LastErr <> "" Then Console.WriteLine(oFS.LastErr)
                        Dim strIn As String = tsIn.ReadAll
                        tsIn.Close()
                        With oEncoder
                            Console.WriteLine("EncodeScriptFile:")
                            Dim strEncData As String = .EncodeScriptFile(".vbs", strIn, 0, "")
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                            Dim oFile As File = oFS.GetFile(strVBSFilePath)
                            Dim strVBEFilePath As String = oFile.ParentFolder.Path & "\" & Replace(oFile.Name, ".vbs", ".vbe")
                            Console.WriteLine("OpenTextFile(" & strVBEFilePath & ")")
                            Dim tsOut As TextStream = oFS.OpenTextFile(strVBEFilePath, FileSystemObject.IOMode.ForWriting, True)
                            If .LastErr <> "" Then Console.WriteLine(.LastErr)
                            tsOut.Write(strEncData)
                            tsOut.Close()
                            Console.WriteLine("OK")
                        End With
                    End If
                    Console.WriteLine("*******************")
            End Select

        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
