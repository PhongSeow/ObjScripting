'**********************************
'* Name: Folder
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.Folder
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 27/2/2021
'**********************************
Public Class Folder
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public Property Attributes() As FileAttribute
		Get
			Try
				Return Me.Obj.Attributes
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As FileAttribute)
			Try
				Me.Obj.Attributes = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Attributes.Set", ex)
			End Try
		End Set
	End Property
	Public Sub Copy(Destination As String, Optional OverWriteFiles As Boolean = True)
		Try
			Me.Obj.Copy(Destination, OverWriteFiles)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Copy", ex)
		End Try
	End Sub
	Public Function CreateTextFile(FileName As String, Optional Overwrite As Boolean = True, Optional Unicode As Boolean = False) As TextStream
		Try
			CreateTextFile = Me.Obj.CreateTextFile(FileName, Overwrite, Unicode)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CreateTextFile", ex)
			Return Nothing
		End Try
	End Function
	Public ReadOnly Property DateCreated() As Date
		Get
			Try
				Return Me.Obj.DateCreated
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DateCreated.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property DateLastAccessed() As Date
		Get
			Try
				Return Me.Obj.DateLastAccessed
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DateLastAccessed.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property DateLastModified() As Date
		Get
			Try
				Return Me.Obj.DateLastModified
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DateLastModified.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Sub Delete(Optional Force As Boolean = False)
		Try
			Me.Obj.Delete(Force = False)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Delete", ex)
		End Try
	End Sub
	Public ReadOnly Property Drive() As Drive
		Get
			Try
				Return Me.Obj.Drive
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Drive.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Files() As Files
		Get
			Try
				Return Me.Obj.Files
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Files.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property IsRootFolder() As Boolean
		Get
			Try
				Return Me.Obj.IsRootFolder
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("IsRootFolder.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Sub Move(Destination As String)
		Try
			Me.Obj.Move(Destination)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Move", ex)
		End Try
	End Sub
	Public Property Name() As String
		Get
			Try
				Return Me.Obj.Name
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Name.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Name = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Name.Set", ex)
			End Try
		End Set
	End Property
	Public ReadOnly Property ParentFolder() As Folder
		Get
			Try
				Return Me.Obj.ParentFolder
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ParentFolder.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Path() As String
		Get
			Try
				Return Me.Obj.Path
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Path.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property ShortName() As String
		Get
			Try
				Return Me.Obj.ShortName
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ShortName.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property ShortPath() As String
		Get
			Try
				Return Me.Obj.ShortPath
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ShortPath.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Size() As Object
		Get
			Try
				Return Me.Obj.Size
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Size.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property SubFolders() As Folders
		Get
			Try
				Return Me.Obj.SubFolders
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("SubFolders.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Type() As String
		Get
			Try
				Return Me.Obj.Type
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Type.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
End Class
