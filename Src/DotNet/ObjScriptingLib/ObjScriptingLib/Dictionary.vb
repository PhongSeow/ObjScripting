'**********************************
'* Name: Dictionary
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.Dictionary
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 3/3/2021
'**********************************
Public Class Dictionary
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Enum CompareMethod
		BinaryCompare = 0
		DatabaseCompare = 2
		TextCompare = 1
	End Enum

	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public Sub Add(Key As String, Item As Object)
		Try
			Me.Obj.Add(Key, Item)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Add", ex)
		End Try
	End Sub
	Public Property CompareMode() As CompareMethod
		Get
			Try
				Return Me.Obj.CompareMode
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CompareMode.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As CompareMethod)
			Try
				Me.Obj.CompareMode = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("CompareMode.Set", ex)
			End Try
		End Set
	End Property
	Public Property Count() As Long
		Get
			Try
				Return Me.Obj.Count
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Count.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Long)
			Try
				Me.Obj.Count = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Count.Set", ex)
			End Try
		End Set
	End Property
	Public Function Exists(Key As String) As Boolean
		Try
			Exists = Me.Obj.Exists(Key)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Exists", ex)
			Return Nothing
		End Try
	End Function
	Public Property Item() As Object
		Get
			Try
				Return Me.Obj.Item
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Item.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As Object)
			Try
				Me.Obj.Item = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Item.Set", ex)
			End Try
		End Set
	End Property
	Public Function Items() As Object
		Try
			Items = Me.Obj.Items
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Items", ex)
			Return Nothing
		End Try
	End Function
	Public Property Key() As String
		Get
			Try
				Return Me.Obj.Key
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Key.Get", ex)
				Return Nothing
			End Try
		End Get
		Set(value As String)
			Try
				Me.Obj.Key = value
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Key.Set", ex)
			End Try
		End Set
	End Property
	Public Function Keys() As Object
		Try
			Keys = Me.Obj.Keys
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Keys", ex)
			Return Nothing
		End Try
	End Function
	Public Sub Remove(Key)
		Try
			Me.Obj.Remove(Key)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Remove", ex)
		End Try
	End Sub
	Public Sub RemoveAll()
		Try
			Me.Obj.RemoveAll
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("RemoveAll", ex)
		End Try
	End Sub
End Class
