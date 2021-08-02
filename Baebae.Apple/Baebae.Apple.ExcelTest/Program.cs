using Baebae.Apple.NPOI;
using Microsoft.Data.Analysis;
using System;
using System.IO;

namespace Baebae.Apple.ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            var filePath = @"";
            var sheetname = "1-2";

            var dataTable = NPOIHelper.ExcelFileToDataTable(filePath,sheetname);

            var df1 = DataFrame.LoadCsv("ohlcdata.csv");

            Console.ReadKey();
        }
    }
}
