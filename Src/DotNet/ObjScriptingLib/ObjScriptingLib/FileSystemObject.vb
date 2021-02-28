'**********************************
'* Name: FileSystemObject
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.FileSystemObject
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 27/2/2021
'**********************************
Public Class FileSystemObject
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Enum SpecialFolderConst
		SystemFolder = 1
		TemporaryFolder = 2
		WindowsFolder = 0
	End Enum
	Public Enum Tristate
		TristateFalse = 0
		TristateMixed = -2
		TristateTrue = -1
		TristateUseDefault = -2
	End Enum
	Public Enum IOMode
		ForAppending = 8
		ForReading = 1
		ForWriting = 2
	End Enum
	Public Enum StandardStreamTypes
		StdErr = 2
		StdIn = 0
		StdOut = 1
	End Enum

	Public Sub New()
		MyBase.New(CLS_VERSION)
		Me.Obj = CreateObject("Scripting.FileSystemObject")
	End Sub
	Public Function BuildPath(Path As String, Name As String) As String
		Try
			BuildPath = Me.Obj.BuildPath(Path, Name)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("BuildPath", ex)
			Return Nothing
		End Try
	End Function
	Public Sub CopyFile(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
		Try
			Me.Obj.CopyFile(Source, Destination, OverWriteFiles)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CopyFile", ex)
		End Try
	End Sub
	Public Sub CopyFolder(Source As String, Destination As String, Optional OverWriteFiles As Boolean = True)
		Try
			Me.Obj.CopyFolder(Source, Destination, OverWriteFiles)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CopyFolder", ex)
		End Try
	End Sub
	Public Function CreateFolder(Path As String) As Folder
		Try
			Dim oFolder As New Folder
			oFolder.Obj = Me.Obj.CreateFolder(Path)
			Return oFolder
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CreateFolder", ex)
			Return Nothing
		End Try
	End Function
	Public Function CreateTextFile(FileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As TextStream
		Try
			Dim oTextStream As New TextStream
			oTextStream.Obj = Me.Obj.CreateTextFile(FileName, Overwrite, Unicode)
			Return oTextStream
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CreateTextFile", ex)
			Return Nothing
		End Try
	End Function
	Public Sub DeleteFile(FileSpec As String, Optional Force As Boolean = False)
		Try
			Me.Obj.DeleteFile(FileSpec, Force)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("DeleteFile", ex)
		End Try
	End Sub
	Public Sub DeleteFolder(FolderSpec As String, Optional Force As Boolean = False)
		Try
			Me.Obj.DeleteFolder(FolderSpec, Force)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("DeleteFolder", ex)
		End Try
	End Sub
	Public Function DriveExists(DriveSpec As String) As Boolean
		Try
			DriveExists = Me.Obj.DriveExists(DriveSpec)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("DriveExists", ex)
			Return Nothing
		End Try
	End Function
	Public ReadOnly Property Drives() As Drives
		Get
			Try
				Return Me.Obj.Drives
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Drives.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function FileExists(FileSpec As String) As Boolean
		Try
			FileExists = Me.Obj.FileExists(FileSpec)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("FileExists", ex)
			Return Nothing
		End Try
	End Function
	Public Function FolderExists(FolderSpec As String) As Boolean
		Try
			FolderExists = Me.Obj.FolderExists(FolderSpec)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("FolderExists", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetAbsolutePathName(Path As String) As String
		Try
			GetAbsolutePathName = Me.Obj.GetAbsolutePathName(Path)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetAbsolutePathName", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetBaseName(Path As String) As String
		Try
			GetBaseName = Me.Obj.GetBaseName(Path)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetBaseName", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetDrive(DriveSpec As String) As Drive
		Try
			Dim oDrive As New Drive
			oDrive.Obj = Me.Obj.GetDrive(DriveSpec)
			Return oDrive
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetDrive", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetDriveName(Path As String) As String
		Try
			GetDriveName = Me.Obj.GetDriveName(Path)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetDriveName", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetExtensionName(Path As String) As String
		Try
			GetExtensionName = Me.Obj.GetExtensionName(Path)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetExtensionName", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetFile(FilePath As String) As File
		Try
			Dim oFile As New File
			oFile.Obj = Me.Obj.GetFile(FilePath)
			Return oFile
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetFile", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetFileName(Path As String) As String
		Try
			GetFileName = Me.Obj.GetFileName(Path)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetFileName", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetFileVersion(FileName As String) As String
		Try
			GetFileVersion = Me.Obj.GetFileVersion(FileName)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetFileVersion", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetFolder(FolderPath As String) As Folder
		Try
			Dim oFolder As New Folder
			oFolder.Obj = Me.Obj.GetFolder(FolderPath)
			Return oFolder
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetFolder", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetParentFolderName(Path As String) As String
		Try
			GetParentFolderName = Me.Obj.GetParentFolderName(Path)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetParentFolderName", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetSpecialFolder(SpecialFolder As SpecialFolderConst) As Folder
		Try
			GetSpecialFolder = Me.Obj.GetSpecialFolder(SpecialFolder)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetSpecialFolder", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetStandardStream(StandardStreamType As StandardStreamTypes, Optional Unicode As Boolean = False) As TextStream
		Try
			GetStandardStream = Me.Obj.GetStandardStream(StandardStreamType, Unicode)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetStandardStream", ex)
			Return Nothing
		End Try
	End Function
	Public Function GetTempName() As String
		Try
			GetTempName = Me.Obj.GetTempName()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("GetTempName", ex)
			Return Nothing
		End Try
	End Function
	Public Sub MoveFile(Source As String, Destination As String)
		Try
			Me.Obj.MoveFile(Source, Destination)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveFile", ex)
		End Try
	End Sub
	Public Sub MoveFolder(Source As String, Destination As String)
		Try
			Me.Obj.MoveFolder(Source, Destination)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("MoveFolder", ex)
		End Try
	End Sub
	Public Function OpenTextFile(FileName As String, Optional IOMode As IOMode = IOMode.ForReading, Optional Create As Boolean = False, Optional Format As Tristate = Tristate.TristateFalse) As TextStream
		Try
			Dim oTextStream As New TextStream
			oTextStream.Obj = Me.Obj.OpenTextFile(FileName, IOMode, Create, Format)
			Return oTextStream
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("OpenTextFile", ex)
			Return Nothing
		End Try
	End Function
End Class
