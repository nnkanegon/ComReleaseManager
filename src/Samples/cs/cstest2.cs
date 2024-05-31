using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Windows.Forms;

/// <summary>
/// Tests for ComReleaseManager (Method Chain style)
/// </summary>
class C1
{
	public static void Test2(dynamic objSheet)
	{
		using (ComReleaseManager crm = new ComReleaseManager())
		{
			var a1value = ((dynamic)(crm.Assign(objSheet)
				.Evaluate((Func<dynamic, object>)(x => x.Range["A1"]))
				.Evaluate((Func<dynamic, object>)(x => x.Offset[0, 0]))
				.Value())).Value;
			((dynamic)(crm.Assign(objSheet)
				.Evaluate((Func<dynamic, object>)(x => x.Range["A1"]))
				.Evaluate((Func<dynamic, object>)(x => x.Offset[2, 2]))
				.Value())).Value = a1value;
		}
	}

	public static void Test1(dynamic objSheet)
	{
		using (ComReleaseManager crm = new ComReleaseManager())
		{
			dynamic objRange1 = crm.Add(objSheet.Range("A1"));
			objRange1.Value = 10;

			dynamic objRange2 = crm.Assign(objSheet)
				.Evaluate((Func<dynamic, object>)(x => x.Cells))
				.Evaluate((Func<dynamic, object>)(x => x.Item[2, 2]))
				.Value();
			objRange2.Value = 11;

			// The following patterns are equivalent.
			// 1)
			// dynamic objRange2 = crm.Assign(objSheet)
			//		.Evaluate((Func<dynamic, object>)(x => x.Cells))
			//		.Evaluate((Func<dynamic, object>)(x => x.Item[2, 2]))
			//		.Value();
			// 2)
			// dynamic objRange2 = crm.Assign(crm.Add(objSheet.Cells))
			//		.Evaluate((Func<dynamic, object>)(x => x.Item[2, 2]))
			//		.Value();
			// 3)
			// dynamic objRange2 = crm.Assign(objSheet.Cells, true)
			//		.Evaluate((Func<dynamic, object>)(x => x.Item[2, 2]))
			//		.Value();
		}
	}

	public static void WriteBook(dynamic objExcel)
	{
		string outputFileName = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "cstest2.xlsx");
		using (ComReleaseManager crm = new ComReleaseManager())
		{
			dynamic objBook = crm.Assign(objExcel)
				.Evaluate((Func<dynamic, object>)(x => x.Workbooks))
				.Evaluate((Func<dynamic, object>)(x => x.Add()))
				.Value();
			try
			{
				using (ComReleaseManager crm2 = new ComReleaseManager())
				{
					dynamic objSheet = crm2.Assign(objBook)
						.Evaluate((Func<dynamic, object>)(x => x.Worksheets))
						.Evaluate((Func<dynamic, object>)(x => x.Item["sheet1"]))
						.Value();

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
