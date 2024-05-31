Imports System.Runtime.InteropServices
Imports System.Collections.Generic

''' <summary>
''' Helper class for releasing COM object
''' </summary>
Public Class ComReleaseManager
	Implements IDisposable

	Dim objList As New List(Of Object)
	Dim currentObject As Object = Nothing
	Dim disposed As Boolean = False

	Public Sub New()
	End Sub

	Public Sub Dispose() Implements IDisposable.Dispose
		Dispose(True)
	End Sub

	Protected Overridable Sub Dispose(disposing As Boolean)
		If disposed Then
			Return
		End If
		if disposing Then
			Me.Release()
		End If
		Me.disposed = True
	End Sub

	Public Function Add(obj As Object) As Object
		objList.Insert(0, obj)
		currentObject = obj
		Return obj
	End Function

	Public Function Assign(obj As Object, Optional isAddObject As Boolean = False) As ComReleaseManager
		If isAddObject Then
			Add(obj)
		End If
		currentObject = obj
		Return Me
    End Function

	Public Function Evaluate(f As Func(Of Object, Object)) As ComReleaseManager
		Add(f(currentObject))
		Return Me
	End Function

	Public Function Value() As Object
		Return currentObject
	End Function

	Private Sub Release()
		For Each obj As Object In objList
			Release(obj)
		Next
		objList.Clear()
		currentObject = Nothing
	End Sub

	Public Shared Sub Release(obj As Object, Optional useFinalRelease As Boolean = False)
		If obj Is Nothing Then
			Return
		End If
		Try
			If useFinalRelease Then
				Marshal.FinalReleaseComObject(obj)
			Else
				Marshal.ReleaseComObject(obj)
			End If
		Catch ex As Exception
		End Try
	End Sub

	Public Shared Sub GCCollect()
		System.GC.Collect()
		System.GC.WaitForPendingFinalizers()
		System.GC.Collect()
	End Sub
End Class
