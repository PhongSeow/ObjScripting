'**********************************
'* Name: Drive
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.Drive
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 27/2/2021
'**********************************
Public Class Drive
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Enum DriveTypeConst
		CDRom = 4
		Fixed = 2
		RamDisk = 5
		Remote = 3
		Removable = 1
		UnknownType = 0
	End Enum
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public ReadOnly Property AvailableSpace() As Object
		Get
			Try
				Return Me.Obj.AvailableSpace
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("AvailableSpace.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property DriveLetter() As String
		Get
			Try
				Return Me.Obj.DriveLetter
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DriveLetter.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property DriveType() As DriveTypeConst
		Get
			Try
				Return Me.Obj.DriveType
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("DriveType.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property FileSystem() As String
		Get
			Try
				Return Me.Obj.FileSystem
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("FileSystem.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property FreeSpace() As Object
		Get
			Try
				Return Me.Obj.FreeSpace
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("FreeSpace.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property IsReady() As Boolean
		Get
			Try
				Return Me.Obj.IsReady
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("IsReady.Get", ex)
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
	Public ReadOnly Property RootFolder() As Folder
		Get
			Try
				Return Me.Obj.RootFolder
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("RootFolder.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property SerialNumber() As Long
		Get
			Try
				Return Me.Obj.SerialNumber
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("SerialNumber.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property ShareName() As String
		Get
			Try
				Return Me.Obj.ShareName
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("ShareName.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property TotalSize() As Object
		Get
			Try
				Return Me.Obj.TotalSize
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("TotalSize.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Property VolumeName() As String
		Get
			Try
				Return Me.Obj.VolumeName
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("VolumeName.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.VolumeName = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("VolumeName.Set", ex)
			End Try
		End Set
	End Property
End Class
