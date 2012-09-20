using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OfficeToXps
{
    class Program
    {
        static void Main(string[] args)
        {
            string SourceDocPath = @"C:\Test\Test.doc";
            string ExportFilePath = @"C:\Test\Test.xps";

            Console.WriteLine("(1) Word");
            Console.WriteLine("(2) Excel");
            Console.WriteLine("(3) PowerPoint");
            Console.WriteLine("(4) Word 2007");
            Console.WriteLine("(5) Excel 2007");
            Console.WriteLine("(6) PowerPoint 2007");
            Console.Write("Enter which app to test: ");

            var selection = Console.ReadKey();
            Console.WriteLine();
            Console.WriteLine("Testing...");
            Console.WriteLine();

            if (selection.Key == ConsoleKey.D2 || selection.Key == ConsoleKey.NumPad2)
            {
                SourceDocPath = Path.ChangeExtension(SourceDocPath, ".xls");
            }
            else if (selection.Key == ConsoleKey.D3 || selection.Key == ConsoleKey.NumPad3)
            {
                SourceDocPath = Path.ChangeExtension(SourceDocPath, ".ppt");
            }
            else if (selection.Key == ConsoleKey.D4 || selection.Key == ConsoleKey.NumPad4)
            {
                SourceDocPath = Path.ChangeExtension(SourceDocPath, ".docx");
            }
            else if (selection.Key == ConsoleKey.D5 || selection.Key == ConsoleKey.NumPad5)
            {
                SourceDocPath = Path.ChangeExtension(SourceDocPath, ".xlsx");
            }
            else if (selection.Key == ConsoleKey.D6 || selection.Key == ConsoleKey.NumPad6)
            {
                SourceDocPath = Path.ChangeExtension(SourceDocPath, ".pptx");
            }

            // This is the stuff for Feature 2
            CallFeature2();

            var result = OfficeToXps.ConvertToXps(SourceDocPath, ref ExportFilePath);

            Console.WriteLine(result.Result.ToString());
            Console.WriteLine();
            Console.WriteLine(result.ResultText);
            Console.WriteLine();

            if (result.ResultError != null)
            {
                Console.WriteLine(result.ResultError.ToString()); 
            }

            Console.WriteLine();
            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }

        private static void CallFeature2()
        {
            Console.WriteLine("Feature 2 stuff here");
        }
    }
}
