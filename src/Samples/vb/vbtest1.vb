Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Windows.Forms

''' <summary>
''' Tests for ComReleaseManager (basic style)
''' </summary>
Module M1
	Public Sub Test2(objSheet As Object)
		Using crm As New ComReleaseManager()

			crm.Add(objSheet.Range("C3")).Value = 
					crm.Add(objSheet.Range("A1")).Value

		End Using
	End Sub

	Public Sub Test1(objSheet As Object)
		Using crm As New ComReleaseManager()

			Dim objRange1 As Object = crm.Add(objSheet.Range("A1"))
			objRange1.Value = 10

			Dim objRange2 As Object = crm.Add(objSheet.Cells)
			Dim objRange3 As Object = crm.Add(objRange2.Item(2, 2))
			objRange3.Value = 11

		End Using
	End Sub

	Public Sub WriteBook(objExcel As Object)
		Dim outputFileName As String = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "vbtest1.xlsx")
		Using crm As New ComReleaseManager()
			Dim objBooks As Object = crm.Add(objExcel.Workbooks)
			Dim objBook As Object = crm.Add(objBooks.Add)
			Try
				Using crm2 As New ComReleaseManager()

					Dim objWorksheets As Object = crm.Add(objBook.Worksheets)
					Dim objSheet As Object = crm.Add(objWorksheets("sheet1"))

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
