using System;
using Wasmtime;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace Catswords.WasmExcel
{
    public class WasmExcel
    {
        // https://github.com/bytecodealliance/wasmtime-dotnet
        static WasmExcel()
        {
            using var engine = new Engine();

            using var module = Module.FromText(
                engine,
                "hello",
                "(module (func $hello (import \"\" \"hello\")) (func (export \"run\") (call $hello)))"
            );

            using var linker = new Linker(engine);
            using var store = new Store(engine);

            linker.Define(
                "",
                "hello",
                Function.FromCallback(store, () => Console.WriteLine("Hello from C#!"))
            );

            var instance = linker.Instantiate(store, module);
            var run = instance.GetFunction(store, "run");
            run?.Invoke(store);
        }

        // https://stackoverflow.com/questions/31038649/passing-an-excel-range-from-vba-to-c-sharp-via-excel-dna
        static Range ReferenceToRange(ExcelReference xlref)
        {
            string refText = (string)XlCall.Excel(XlCall.xlfReftext, xlref, true);
            dynamic app = ExcelDnaUtil.Application;
            return app.Range[refText];
        }

        // https://excel-dna.net/
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        // https://stackoverflow.com/questions/14896215/how-do-you-set-the-value-of-a-cell-using-excel-dna
        [ExcelFunction(Category = "Foo", Description = "Sets value of cell")]
        public static string Foo(String idx)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Range range = app.ActiveCell;

            var dummyData = new object[2, 2] {
                { "foo", "bar" },
                { 2500, 7500 }
            };

            var reference = new ExcelReference(
                range.Row, range.Row + 2 - 1, // from-to-Row
                range.Column - 1, range.Column + 2 - 1); // from-to-Column

            // Cells are written via this async task
            ExcelAsyncUtil.QueueAsMacro(() => { reference.SetValue(dummyData); });

            // Value displayed in the current cell. 
            // It still is a UDF and can be executed multiple times via F2, Return.
            return "=Foo()";
        }
    }
}
