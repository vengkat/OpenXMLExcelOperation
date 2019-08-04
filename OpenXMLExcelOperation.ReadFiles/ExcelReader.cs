using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLExcelOperation.ReadFiles
{
    public class ExcelReader
    {
        #region ReadFileWithText
        /// <summary>
        /// ReadFileWithText - Reads the values of each cell in an excel file
        /// </summary>
        /// <param name="fileName"></param>
        public void ReadFileWithText(string fileName)
        {
            //string fileName = @"C:\Files\Test.xlsx";

            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbook                 = doc.WorkbookPart;
                    SharedStringTablePart StringTablePart = workbook.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable stringTable         = StringTablePart.SharedStringTable;

                    WorksheetPart worksheetPart = workbook.WorksheetParts.First();
                    Worksheet workSheet         = worksheetPart.Worksheet;

                    var rows = workSheet.Descendants<Row>();

                    Console.WriteLine($"Row count = {rows.LongCount()}");

                    // Loop through each row in the workSheet
                    foreach (Row row in rows)
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            // Loop through each cell in the row
                            if (cell.DataType != null)//For number values cell.DataType appears to be null
                            {
                                try
                                {
                                    int cellId = int.Parse(cell.CellValue.Text);
                                    string value = stringTable.ChildElements[cellId].InnerText;
                                    //This prints the text values in the cell
                                    Console.WriteLine($"Cell value: {value}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Internal server error | {ex.Message}");
                                }
                            }
                            else
                            {                     
                                //This prints the numeric values in the cell
                                Console.WriteLine($"Cell value: {cell.CellValue?.Text}");
                            }
                        }
                    }
                }
            }
        }
        #endregion
    }
}
