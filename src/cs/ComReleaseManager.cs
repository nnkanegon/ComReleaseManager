using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

/// <summary>
/// Helper class for releasing COM object
/// </summary>
class ComReleaseManager : IDisposable
{
	private List<object> objList = new List<object>();
	private object currentObject = null;
	private bool disposed = false;

	public ComReleaseManager()
	{
	}

	public void Dispose()
	{
		Dispose(true);
	}

	protected virtual void Dispose(bool disposing)
	{
		if (disposed)
		{
			return;
		}
		if (disposing)
		{
			this.Release();
		}
		this.disposed = true;
	}

	public object Add(object obj)
	{
		objList.Insert(0, obj);
		currentObject = obj;
		return obj;
	}

	public ComReleaseManager Assign(object obj, bool isAddObject = false)
	{
		if (isAddObject)
		{
			Add(obj);
		}
		currentObject = obj;
		return this;
	}

	public ComReleaseManager Evaluate(Func<dynamic, object> func)
	{
		Add(func(currentObject));
		return this;
	}

	public object Value()
	{
		return currentObject;
	}

	private void Release()
	{
		foreach (object obj in objList)
		{
			Release(obj, false);
		}
		objList.Clear();
		currentObject = null;
	}

	public static void Release(object obj, bool useFinalRelease = false)
	{
		if (obj == null)
		{
			return;
		}
		try
		{
			if (useFinalRelease)
			{
				Marshal.FinalReleaseComObject(obj);
			}
			else
			{
				Marshal.ReleaseComObject(obj);
			}
		}
		catch(Exception) {}
	}

	public static void GCCollect()
	{
		System.GC.Collect();
		System.GC.WaitForPendingFinalizers();
		System.GC.Collect();
	}
}
