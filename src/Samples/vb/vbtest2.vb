Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Windows.Forms

''' <summary>
''' Tests for ComReleaseManager (method chain style)
''' </summary>
Module M1
	Public Sub Test2(objSheet As Object)
		Using crm As New ComReleaseManager()

			crm.Add(objSheet.Range("C3")).Value = 
					crm.Add(objSheet.Range("A1")).Value

			Dim  a1value As Object = crm.Assign(objSheet) _
					.Evaluate(Function(x) x.Range("A1")) _
					.Evaluate(Function(x) x.Offset(0, 0)) _
					.Value().Value()
			crm.Assign(objSheet) _
					.Evaluate(Function(x) x.Range("A1")) _
					.Evaluate(Function(x) x.Offset(2, 2)) _
					.Value().Value() = a1value

		End Using
	End Sub

	Public Sub Test1(objSheet As Object)
		Using crm As New ComReleaseManager()

			Dim objRange1 As Object = crm.Add(objSheet.Range("A1"))
			objRange1.Value = 10

			Dim objRange2 As Object = crm.Assign(objSheet) _
					.Evaluate(Function(x) x.Cells) _
					.Evaluate(Function(x) x.Item(2, 2)) _
					.Value()
			objRange2.Value = 11

			'' The following patterns are equivalent.
			'' 1)
			'' Dim objRange2 = crm.Assign(objSheet)
			''		.Evaluate(Function(x) x.Cells) _
			''		.Evaluate(Function(x) x.Item(2, 2)) _
			''		.Value()
			'' 2)
			'' Dim objRange2 = crm.Assign(crm.Add(objSheet.Cells))
			''		.Evaluate(Function(x) x.Item(2, 2)) _
			''		.Value()
			'' 3)
			'' Dim objRange2 = crm.Assign(objSheet.Cells, True)
			''		.Evaluate(Function(x) x.Item(2, 2)) _
			''		.Value()

		End Using
	End Sub

	Public Sub WriteBook(objExcel As Object)
		Dim outputFileName As String = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "vbtest2.xlsx")
		Using crm As New ComReleaseManager()
			Dim objBook As Object = crm.Assign(objExcel) _
					.Evaluate(Function(x) x.Workbooks) _
					.Evaluate(Function(x) x.Add()) _
					.Value()
			Try
				Using crm2 As New ComReleaseManager()

					Dim objSheet As Object = crm.Assign(objBook) _
							.Evaluate(Function(x) x.Worksheets) _
							.Evaluate(Function(x) x.Item("sheet1")) _
							.Value()

					Test1(objSheet)
					Test2(objSheet)

					objExcel.DisplayAlerts = False
					objBook.SaveAs(outputFileName)
					objExcel.DisplayAlerts = True

				End Using
			Finally
				objBook.Close(False)
			End Try

		End Using
	End Sub

	Public Sub WriteExcel()
		Dim objExcel As Object = Nothing
		Try
			objExcel = CreateObject("Excel.Application")
			objExcel.Visible = False
			WriteBook(objExcel)
		Finally
			If objExcel IsNot Nothing Then
				ComReleaseManager.GCCollect()
				objExcel.Quit()
				ComReleaseManager.Release(objExcel)
				objExcel = Nothing
				ComReleaseManager.GCCollect()
			End If
		End Try
	End Sub

	Public Sub Main()
		WriteExcel()
	End Sub
End Module
