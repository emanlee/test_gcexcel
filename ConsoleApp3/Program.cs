using System;
using System.Drawing;
using System.IO;
using GrapeCity.Documents.Excel;
using static System.Net.Mime.MediaTypeNames;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            //create a new workbook
            var workbook = new GrapeCity.Documents.Excel.Workbook();

            IWorksheet sheet = workbook.Worksheets[0];

            //set style.
            sheet.Range["A1"].Value = "Sheet1";
            sheet.Range["A1"].Font.Name = "Wide Latin";
            sheet.Range["A1"].Font.Color = Color.Red;
            sheet.Range["A1"].Interior.Color = Color.Green;

            //change the path to real export path when save.
            sheet.Save(System.IO.Path.Combine(System.Environment.CurrentDirectory, "dest.pdf"), SaveFileFormat.Pdf);
        }
    }
}
