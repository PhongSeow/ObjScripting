'**********************************
'* Name: Encoder
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.Encoder
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 3/3/2021
'**********************************
Public Class Encoder
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
		Me.Obj = CreateObject("Scripting.Encoder")
	End Sub
	Public Function EncodeScriptFile(szExt As String, bstrStreamIn As String, cFlags As Long, bstrDefaultLang As String) As String
		Try
			EncodeScriptFile = Me.Obj.EncodeScriptFile(szExt, bstrStreamIn, cFlags, bstrDefaultLang)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("EncodeScriptFile", ex)
			Return Nothing
		End Try
	End Function
End Class
