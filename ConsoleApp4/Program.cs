using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GrapeCity.Documents.Excel;
using static System.Net.Mime.MediaTypeNames;

namespace ConsoleApp4
{
    class Program
    {
        static void Main(string[] args)
        {

            //create a new workbook
            var workbook = new GrapeCity.Documents.Excel.Workbook();

            Stream fileStream = Application.GetResourceStream("xlsx\\Employee absence schedule.xlsx");
            # Stream fileStream = this.GetResourceStream("xlsx\\Employee absence schedule.xlsx");
            workbook.Open(fileStream);

            //save to a pdf file
            workbook.Save("convertexceltopdf.pdf");


        }
    }
}
