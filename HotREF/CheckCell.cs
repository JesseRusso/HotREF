using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace HotREF
{
    class CheckCell
    {
        private string filePath;
        private SpreadsheetDocument calcSheet;
        public CheckCell(string excelPath)
        {
            using (FileStream fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {

            }
        }
        public string GetCellValue(string sheetName, string refCell)
        {
            string value = null;

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart wbPart = spreadsheet.WorkbookPart;
                    Sheet theSheet = (Sheet)wbPart.Workbook.Descendants<Sheet>().

                        Where(s => s.Name == sheetName).FirstOrDefault();

                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
                    }
                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                    Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == refCell).FirstOrDefault();

                    if (theCell.CellValue != null)
                    {
                        value = theCell.CellValue.InnerText;

                        if (theCell.DataType != null)
                        {
                            switch (theCell.DataType.Value)
                            {
                                case CellValues.SharedString:

                                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                                    if (stringTable != null)
                                    {
                                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                    }
                                    break;

                                case CellValues.Boolean:
                                    switch (value)
                                    {
                                        case "0":
                                            value = "FALSE";
                                            break;
                                        case "1":
                                            value = "TRUE";
                                            break;
                                    }
                                    break;

                                case CellValues.String:
                                    value = theCell.LastChild.InnerText;
                                    break;
                            }
                        }
                    }
                    else
                    {
                        value = null;
                    }
                }
            }
            return value;
        }
    }
}
