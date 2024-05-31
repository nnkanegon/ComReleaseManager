# ComReleaseManager

## 概要

C# や VB.NET で Excel などの COM 操作を行う際、取得したオブジェクトを必ず解放しなければならず面倒です。ComReleaseManager はその面倒を軽減する簡単な解放ヘルパクラスです。

仕組みは単純です。ComReleaseManager のインスタンスを作成後、複数の COM オブジェクトを登録することができ、Dispose メソッドを呼ぶことで、登録したオブジェクトを登録時とは逆の順番でまとめて解放できます。ComReleaseManager は IDisposable を継承するため、using ブロックとともに使用することで、ブロック終端での解放処理が保証されます。

COM オブジェクトの解放が面倒なのは、オブジェクトを作成するタイミングと解放するタイミングが異なるためです。解放ヘルパを使用することで、作成した COM オブジェクトを作成のタイミングで解放ヘルパに登録することができ、登録したオブジェクトはブロックの終端で自動的に解放されます。明示的な解放処理を記述する必要がなくなりソースコード自体も簡素化できます。

## 使用例

簡単な使い方の例を示します。

```cs:test.cs
using System;
using System.Runtime.InteropServices;

class C1
{
    // 基本形
    public static void foo(dynamic objExcel)
    {
        // using ブロック付きで解放ヘルパのインスタンスを作成する
        using (ComReleaseManager crm = new ComReleaseManager())
        {
            // オブジェクトを取得したら即座に解放ヘルパクラスに登録する
            // Add() メソッドは登録したオブジェクト自体を返すので、戻り値をそのまま変数に代入できる
            dynamic objBooks = crm.Add(objExcel.WorkBooks);
            dynamic objBook = crm.Add(objBooks.Add());
            dynamic objSheets = crm.Add(objBook.Worksheets);
            dynamic objSheet = crm.Add(objSheets.Item["sheet1"]);
            dynamic objRange = crm.Add(objSheet.Range["A1"]);
            objRange.Value = 123;

            objBook.Close(false);
        }
        // 登録したオブジェクトはこのタイミングで自動的に解放される
    }

    // メソッドチェーン風
    public static void bar(dynamic objExcel)
    {
        // using ブロック付きで解放ヘルパのインスタンスを作成する
        using (ComReleaseManager crm = new ComReleaseManager())
        {
            // ベースのオブジェクトを Assign() で設定後、Evaluate() メソッドでつなげていき、
            // 最後に Value() でオブジェクトを取得する
            // 中間の一時オブジェクトは Evaluate() メソッド内で解放対象として登録される
            dynamic objBook = crm.Assign(objExcel)
                    .Evaluate((Func<dynamic, object>)(x => x.Workbooks))
                    .Evaluate((Func<dynamic, object>)(x => x.Add()))
                    .Value();
            dynamic objRange = crm.Assign(objBook)
                    .Evaluate((Func<dynamic, object>)(x => x.Worksheets))
                    .Evaluate((Func<dynamic, object>)(x => x.Item["sheet1"]))
                    .Evaluate((Func<dynamic, object>)(x => x.Range["C3"]))
                    .Value();
            objRange.Value = 123;

            objBook.Close(false);
        }
        // 登録したオブジェクトはこのタイミングで自動的に解放される
    }

    public static void Main(string[] args)
    {
        dynamic objExcel = null;
        try
        {
            objExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));

            foo(objExcel);
            bar(objExcel);
        }
        finally
        {
            if (objExcel != null)
            {
                ComReleaseManager.GCCollect();
                objExcel.Quit();
                ComReleaseManager.Release((object)objExcel);
                objExcel = null;
                ComReleaseManager.GCCollect();
            }
        }
    }
}
```

```vb:test.vb
Imports System.Runtime.InteropServices
Imports System.Reflection

Module M1
    ' 基本形
    Public Sub Foo(objExcel As Object)
        ' using ブロック付きで解放ヘルパのインスタンスを作成する
        Using crm As New ComReleaseManager()
            ' オブジェクトを取得したら即座に解放ヘルパクラスに登録する
            ' Add() メソッドは登録したオブジェクト自体を返すので、戻り値をそのまま変数に代入できる
            Dim objBooks As Object = crm.Add(objExcel.Workbooks)
            Dim objBook As Object = crm.Add(objBooks.Add)
            Dim objSheets As Object = crm.Add(objBook.Worksheets)
            Dim objSheet As Object = crm.Add(objSheets("sheet1"))
            Dim objRange As Object = crm.Add(objSheet.Range("A1"))
            objRange.Value = 123

            objBook.Close(False)
        End Using
        ' 登録したオブジェクトはこのタイミングで自動的に解放される
    End Sub

    ' メソッドチェーン風
    Public Sub Bar(objExcel As Object)
        ' using ブロック付きで解放ヘルパのインスタンスを作成する
        Using crm As New ComReleaseManager()
            ' ベースのオブジェクトを Assign() で設定後、Evaluate() メソッドでつなげていき、
            ' 最後に Value() でオブジェクトを取得する
            ' 中間の一時オブジェクトは Evaluate() メソッド内で解放対象として登録される
            Dim objBook As Object = crm.Assign(objExcel) _
                    .Evaluate(Function(x) x.Workbooks) _
                    .Evaluate(Function(x) x.Add()) _
                    .Value()
            Dim objRange As Object = crm.Assign(objBook) _
                    .Evaluate(Function(x) x.Worksheets) _
                    .Evaluate(Function(x) x.Item("sheet1")) _
                    .Evaluate(Function(x) x.Range("C3")) _
                    .Value()
            objRange.Value = 123

            objBook.Close(False)
        End Using
        ' 登録したオブジェクトはこのタイミングで自動的に解放される
    End Sub

    Public Sub Main()
        Dim objExcel As Object = Nothing
        Try
            objExcel = CreateObject("Excel.Application")
            Foo(objExcel)
            Bar(objExcel)
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
End Module
```

上記サンプルでは、Excel オブジェクトに対しては解放ヘルパを使用していません。
万が一例外が発生した場合でも Excel オブジェクトが Quit() されるように finally ブロック内で後処理を行っていますが、このような try-finally を用いるパターンは、using とは相性が悪く個別に対応する必要があります。実際のアプリでは Book も同じように処理した方がよいかもしれません。

なお、解放オブジェクトの登録を代入式の右辺や左辺に書くこともできます。以下は左辺に書いた例です。無理やりですが。

```cs
((dynamic)(crm.Add(objRange.Item[2, 2]))).value = 11;
((dynamic)(crm.Assign(objSheet)
    .Evaluate((Func<dynamic, object>)(x => x.Range["e5"]))
    .Evaluate((Func<dynamic, object>)(x => x.Offset[2, 2]))
    .Value())).value = 22;
```

