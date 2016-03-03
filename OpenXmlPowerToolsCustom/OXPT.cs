using System;
using OpenXmlPowerTools;
using System.Collections.Generic;

namespace Main
{
    /// <summary>
    /// Use Open-Xml-PowerTools custom
    /// https://github.com/shimitei/Open-Xml-PowerTools
    /// </summary>
    class OXPT : IExcelReaderWriter
    {
        public const string Id = "OpenXmlPowerTools";
        public string GetName()
        {
            return Id;
        }

        public string Read(string inFile)
        {
            // no support
            throw new NotImplementedException();
        }

        public void Write(TestTable table, string outFile)
        {
            var ws = new WorksheetDfn
            {
                Name = "data",
                Cols = new ColDfn[]
                {
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                    new ColDfn { AutoFit = new ColAutoFit() },
                },
                Rows = GetRowsEnum(table),
            };


            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    ws,
                },
            };
            SpreadsheetWriter.Write(outFile, wb);
        }

        IEnumerable<RowDfn> GetRowsEnum(TestTable table)
        {
            var border = CellStyleBorder.CreateBoxBorder(CellStyleBorder.Thin);
            var defaultCellStyle = new CellStyleDfn
            {
                Border = border,
            };
            var dateCellStyle = new CellStyleDfn
            {
                Border = border,
                NumFmt = new CellStyleNumFmt { formatCode = "yyyy/mm/dd" },
            };
            int rowCount = table.RowCount;
            var rows = new List<RowDfn>();
            for (int r = 0; r < rowCount; r++)
            {
                var datarow = table.GetNextRow();
                var cells = new CellDfn[]
                {
                    new CellDfn { CellDataType = CellDataType.String, Style = defaultCellStyle, Value = datarow.Col1 },
                    new CellDfn { CellDataType = CellDataType.String, Style = defaultCellStyle, Value = datarow.Col2 },
                    new CellDfn { CellDataType = CellDataType.String, Style = defaultCellStyle, Value = datarow.Col3 },
                    new CellDfn { CellDataType = CellDataType.String, Style = defaultCellStyle, Value = datarow.Col4 },
                    new CellDfn { CellDataType = CellDataType.String, Style = defaultCellStyle, Value = datarow.Col5 },
                    new CellDfn { CellDataType = CellDataType.Number, Style = defaultCellStyle, Value = datarow.Col6 },
                    new CellDfn { CellDataType = CellDataType.Number, Style = defaultCellStyle, Value = datarow.Col7 },
                    new CellDfn { CellDataType = CellDataType.Number, Style = defaultCellStyle, Value = datarow.Col8 },
                    new CellDfn { CellDataType = CellDataType.Number, Style = defaultCellStyle, Value = datarow.Col9 },
                    new CellDfn { CellDataType = CellDataType.Number, Style = defaultCellStyle, Value = datarow.Col10 },
                    new CellDfn { CellDataType = CellDataType.Date, Style = dateCellStyle, Value = datarow.Col11 },
                    new CellDfn { CellDataType = CellDataType.Date, Style = dateCellStyle, Value = datarow.Col12 },
                    new CellDfn { CellDataType = CellDataType.Date, Style = dateCellStyle, Value = datarow.Col13 },
                    new CellDfn { CellDataType = CellDataType.Date, Style = dateCellStyle, Value = datarow.Col14 },
                    new CellDfn { CellDataType = CellDataType.Date, Style = dateCellStyle, Value = datarow.Col15 },
                };
                var row = new RowDfn { Cells = cells };
                //yield return row;
                rows.Add(row);
            }
            return rows.ToArray();
        }
    }
}
