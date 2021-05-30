using Newtonsoft.Json;
using System;
using System.Collections;
using System.Numerics;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
//using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
//using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Expressions;
using GrapeCity.Documents.Excel.Drawing;
using System.Globalization;
using System.IO;

namespace GrapeCity.Documents.Excel.Examples.Features.PDFExporting.ConvertExcelToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
					//create a new workbook
		var workbook = new Workbook();
		
		Stream fileStream = GetResourceStream("xlsx\\№ЬАн.xlsx");
		workbook.Open(fileStream);
		        
		//save to a pdf file
		workbook.Save("convertexceltopdf.pdf");

        }

		static Stream GetResourceStream(string resourcePath)
        {
            string resource = "ConvertExcelToPDF.Resource." + resourcePath.Replace("\\", ".");
            var assembly = typeof(Program).GetTypeInfo().Assembly;
            return assembly.GetManifestResourceStream(resource);
        }

        static Stream GetResourceStream2(string resourcePath)
        {
            string resource = "ConvertExcelToPDF.Resource." + resourcePath.Replace("\\", ".");
            var assembly = typeof(Program).GetTypeInfo().Assembly;
            return assembly.GetManifestResourceStream(resource);
        }


    }
}