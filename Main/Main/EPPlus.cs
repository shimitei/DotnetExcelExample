using OfficeOpenXml.Style;
using System.IO;
using System.Text;

namespace Main
{
    class EPPlus : IExcelReaderWriter
    {
        public const string Id = "EPPlus";
        public string GetName()
        {
            return Id;
        }

        public string Read(string inFile)
        {
            var newFile = new FileInfo(inFile);
            var pck = new OfficeOpenXml.ExcelPackage(newFile);
            var sheet = pck.Workbook.Worksheets[1];

            var sb = new StringBuilder();
            int rows = sheet.Dimension.Rows;
            int cols = sheet.Dimension.Columns;

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    sb.Append(sheet.Cells[r, c].Text.Replace("\n", "") + "\t");
                }
                sb.Append("\r\n");
            }
            return sb.ToString();
        }

        public void Write(TestTable table, string outFile)
        {
            var pck = new OfficeOpenXml.ExcelPackage();
            var sheet = pck.Workbook.Worksheets.Add("data");

            const string dateFormat = "yyyy/mm/dd";
            for (int r = 1; r <= table.RowCount; r++)
            {
                var datarow = table.GetNextRow();
                OfficeOpenXml.ExcelRange cell;
                var colNum = 1;
                //STRING-column
                //col1
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col1;
                colNum++;
                //col2
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col2;
                colNum++;
                //col3
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col3;
                colNum++;
                //col4
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col4;
                colNum++;
                //col5
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col5;
                colNum++;
                //NUMBER-column
                //col6
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col6;
                colNum++;
                //col7
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col7;
                colNum++;
                //col8
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col8;
                colNum++;
                //col9
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col9;
                colNum++;
                //col10
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col10;
                colNum++;
                //DATE-column
                //col11
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col11;
                cell.Style.Numberformat.Format = dateFormat;
                colNum++;
                //col12
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col12;
                cell.Style.Numberformat.Format = dateFormat;
                colNum++;
                //col13
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col13;
                cell.Style.Numberformat.Format = dateFormat;
                colNum++;
                //col14
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col14;
                cell.Style.Numberformat.Format = dateFormat;
                colNum++;
                //col15
                cell = sheet.Cells[r, colNum];
                cell.Value = datarow.Col15;
                cell.Style.Numberformat.Format = dateFormat;
                colNum++;
            }

            //AutoFit
            for (var i = 1; i <= table.ColCount; i++)
            {
                sheet.Column(i).AutoFit();
            }
            //border-style
            {
                var border = sheet.Cells[1, 1, table.RowCount, table.ColCount].Style.Border;
                border.Top.Style = ExcelBorderStyle.Thin;
                border.Left.Style = ExcelBorderStyle.Thin;
                border.Bottom.Style = ExcelBorderStyle.Thin;
                border.Right.Style = ExcelBorderStyle.Thin;
            }

            FileInfo outFileInfo = new FileInfo(outFile);
            pck.SaveAs(outFileInfo);
        }
    }
}
