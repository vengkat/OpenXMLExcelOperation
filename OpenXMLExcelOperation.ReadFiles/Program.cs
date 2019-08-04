using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLExcelOperation.ReadFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName    = @"C:\Files\Test.xlsx";
            //Or read file path from console
            if (File.Exists(fileName))
            {
                try
                {
                    ExcelReader reader = new ExcelReader();
                    reader.ReadFileWithText(fileName);
                    Console.ReadLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Internal server error | {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Invalid file path specified!");
            }
        }
    }
}
