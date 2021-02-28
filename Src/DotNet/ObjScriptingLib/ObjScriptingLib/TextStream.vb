'**********************************
'* Name: TextStream
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Mapping VB6 Scripting.TextStream
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.1
'* Create Time: 27/2/2021
'**********************************
Public Class TextStream

	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.1"
	Public Obj As Object
	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub
	Public ReadOnly Property AtEndOfLine() As Boolean
		Get
			Try
				Return Me.Obj.AtEndOfLine
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("AtEndOfLine.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property AtEndOfStream() As Boolean
		Get
			Try
				Return Me.Obj.AtEndOfStream
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("AtEndOfStream.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Sub Close()
		Try
			Me.Obj.Close()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Close", ex)
		End Try
	End Sub
	Public ReadOnly Property Column() As Long
		Get
			Try
				Return Me.Obj.Column
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Column.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public ReadOnly Property Line() As Long
		Get
			Try
				Return Me.Obj.Line
				Me.ClearErr()
			Catch ex As Exception
				Me.SetSubErrInf("Line.Get", ex)
				Return Nothing
			End Try
		End Get
	End Property
	Public Function Read(Characters As Long) As String
		Try
			Read = Me.Obj.Read(Characters)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Read", ex)
			Return Nothing
		End Try
	End Function
	Public Function ReadAll() As String
		Try
			ReadAll = Me.Obj.ReadAll()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("ReadAll", ex)
			Return Nothing
		End Try
	End Function
	Public Function ReadLine() As String
		Try
			ReadLine = Me.Obj.ReadLine()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("ReadLine", ex)
			Return Nothing
		End Try
	End Function
	Public Sub Skip(Characters As Long)
		Try
			Me.Obj.Skip(Characters)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Skip", ex)
		End Try
	End Sub
	Public Sub SkipLine()
		Try
			Me.Obj.SkipLine()
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("SkipLine", ex)
		End Try
	End Sub
	Public Sub Write(Text As String)
		Try
			Me.Obj.Write(Text)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("Write", ex)
		End Try
	End Sub
	Public Sub WriteBlankLines(Lines As Long)
		Try
			Me.Obj.WriteBlankLines(Lines)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("WriteBlankLines", ex)
		End Try
	End Sub
	Public Sub WriteLine(Optional Text As String = "")
		Try
			Me.Obj.WriteLine(Text)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("WriteLine", ex)
		End Try
	End Sub
End Class
