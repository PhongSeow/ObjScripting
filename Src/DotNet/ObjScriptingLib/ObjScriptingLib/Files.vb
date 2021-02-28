﻿'**********************************
'* Name: Files
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.Files
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 27/2/2021
'**********************************
Public Class Files
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public ReadOnly Property Count() As Long
		Get
			Try
				Return Me.Obj.Count
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Count.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Item(Key) As File
		Get
			Try
				Return Me.Obj.Item(Key)
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Item.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
End Class
