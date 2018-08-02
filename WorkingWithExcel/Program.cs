using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkingWithExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\Users\waqas.dilawar\Documents\demodocument.xlsx";
            string value = string.Empty;
            value = XLGetValue(fileName, "Sheet1");
            Console.WriteLine(value);
            Console.ReadLine();
        }
        static string XLGetValue(string fileName, string sheetName)
        {
            string value = null;
            var error = string.Empty;
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                      FirstOrDefault(s => s.Name == sheetName);

                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
                    }
                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                    Worksheet workSheet = wsPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Descendants<Row>();
                    var rowss = sheetData.Elements<Row>();

                    //skip first heading row and check for non-empty rows
                    foreach (var row in rows.Skip(1).Where(d => d.InnerText != ""))
                    {
                        var theRow = rows.FirstOrDefault(r => r.RowIndex.Value == row.RowIndex.Value);

                        //get only 2 (A and B) columns, if require others then take accordingly
                        foreach (Cell cell in theRow.Take(2))
                        {
                            if (cell.DataType != null)
                            {
                                switch (cell.DataType.Value)
                                {
                                    case CellValues.SharedString:
                                        var stringTable = wbPart.SharedStringTablePart;
                                        if (stringTable != null)
                                        {
                                            var textItem = stringTable.SharedStringTable.
                                                ElementAtOrDefault(int.Parse(cell.InnerText));
                                            if (textItem != null)
                                            {
                                               

                                                if (cell.CellReference.ToString().Contains("A"))
                                                {
                                                    value +="Cell Value = "+ textItem.InnerText + " of cell = " + cell.CellReference.ToString();
                                                }
                                            }
                                        }
                                        break;

                                    case CellValues.Boolean:
                                        switch (value)
                                        {
                                            case "0":
                                                value = "FALSE";
                                                break;
                                            default:
                                                value = "TRUE";
                                                break;
                                        }
                                        break;

                                }
                            }
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                error += "Error occured<br/>";
            }

            return value;
        }
    }
}
