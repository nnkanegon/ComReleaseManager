using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;

/// <summary>
/// Tests for ComReleaseManager (basic style)
/// </summary>
class C1
{
	public static void Test2(dynamic objSheet)
	{
		using (ComReleaseManager crm = new ComReleaseManager())
		{
			((dynamic)(crm.Add(objSheet.Range("C3")))).Value = 
					((dynamic)(crm.Add(objSheet.Range("A1")))).Value;
		}
	}

	public static void Test1(dynamic objSheet)
	{
		using (ComReleaseManager crm = new ComReleaseManager())
		{
			dynamic objRange1 = crm.Add(objSheet.Range("A1"));
			objRange1.Value = 10;

			dynamic objRange2 = crm.Add(objSheet.Cells);
			dynamic objRange3 = crm.Add(objRange2.Item[2, 2]);
			objRange3.Value = 11;
		}
	}

	public static void WriteBook(dynamic objExcel)
	{
		string outputFileName = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "cstest1.xlsx");
		using (ComReleaseManager crm = new ComReleaseManager())
		{
			dynamic objWorkbooks = crm.Add(objExcel.Workbooks);
			dynamic objBook = crm.Add(objWorkbooks.Add());
			try
			{
				using (ComReleaseManager crm2 = new ComReleaseManager())
				{
					dynamic objWorksheets = crm2.Add(objBook.Worksheets);
					dynamic objSheet = crm2.Add(objWorksheets.Item["sheet1"]);

					Test1(objSheet);
					Test2(objSheet);

					objExcel.DisplayAlerts = false;
					objBook.SaveAs(outputFileName);
					objExcel.DisplayAlerts = true;
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			finally
			{
				objBook.Close(false);
			}
		}
	}

	public static void WriteExcel()
	{
		dynamic objExcel = null;
		try
		{
			objExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
			objExcel.Visible = false;
			WriteBook(objExcel);
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex.Message);
		}
		finally
		{
			if (objExcel != null) {
				ComReleaseManager.GCCollect();
				objExcel.Quit();
				ComReleaseManager.Release(objExcel);
				objExcel = null;
				ComReleaseManager.GCCollect();
			}
		}
	}

	public static void Main(string[] args)
	{
		WriteExcel();
	}
}
