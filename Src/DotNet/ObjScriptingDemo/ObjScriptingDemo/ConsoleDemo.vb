﻿Imports ObjScriptingLib

Public Class ConsoleDemo
    Private oFS As New FileSystemObject

    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to FileSystemObject")
            Console.WriteLine("Press B to TextStream")
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
                        Console.WriteLine("*******************")
                        Select Case Console.ReadKey().Key
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path")
                                strFilePath = Console.ReadLine()
                                Dim oFile As File = oFS.GetFile(strFilePath)
                                Console.WriteLine("DateCreated: " & oFile.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFile.DateLastModified)
                                Console.WriteLine("Name: " & oFile.Name)
                                Console.WriteLine("Path: " & oFile.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.B
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file folder ")
                                strFolderPath = Console.ReadLine()
                                Dim oFolder As Folder = oFS.GetFolder(strFolderPath)
                                Console.WriteLine("DateCreated: " & oFolder.DateCreated)
                                Console.WriteLine("DateLastModified: " & oFolder.DateLastModified)
                                Console.WriteLine("Name: " & oFolder.Name)
                                Console.WriteLine("Path: " & oFolder.Path)
                                Console.WriteLine("#################")
                            Case ConsoleKey.C
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path")
                                strFilePath = Console.ReadLine()
                                Console.WriteLine("FileExists: " & oFS.FileExists(strFilePath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.D
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file folder")
                                strFolderPath = Console.ReadLine()
                                Console.WriteLine("FolderExists: " & oFS.FolderExists(strFolderPath))
                                Console.WriteLine("#################")
                            Case ConsoleKey.E
                                Dim strFolderPath As String
                                Console.WriteLine("#################")
                                With oFS
                                    Console.WriteLine("Enter the file folder")
                                End With
                                strFolderPath = Console.ReadLine()
                                oFS.CreateFolder(strFolderPath)
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
                                Dim strFilePath As String
                                Console.WriteLine("#################")
                                Console.WriteLine("Enter the file path")
                                strFilePath = Console.ReadLine()
                                If oFS.FileExists(strFilePath) = False Then
                                    Console.WriteLine(strFilePath & " not found.")
                                Else
                                    Dim oTextStream As TextStream
                                    Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                    oTextStream = oFS.OpenTextFile(strFilePath, FileSystemObject.IOMode.ForReading, False)
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
                                Dim strFilePath As String
                                Console.Write("Enter the file path:")
                                strFilePath = Console.ReadLine()
                                Dim oTextStream As TextStream
                                Console.WriteLine("OpenTextFile(" & strFilePath & ")...")
                                oTextStream = oFS.OpenTextFile(strFilePath, FileSystemObject.IOMode.ForWriting, True)
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
            End Select

        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
