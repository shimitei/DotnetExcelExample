using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;

namespace Main
{
    class NPOI : IExcelReaderWriter
    {
        public const string Id = "NPOI";
        public string GetName()
        {
            return Id;
        }

        public string Read(string inFile)
        {
            var stream = File.OpenRead(inFile);
            var book = new XSSFWorkbook(stream);
            stream.Close();

            var sb = new StringBuilder();
            var sheet = book.GetSheetAt(0);
            int lastRowNum = sheet.LastRowNum;
            for (int r = 0; r <= lastRowNum; r++)
            {
                var datarow = sheet.GetRow(r);
                {
                    foreach (var cell in datarow.Cells)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.Numeric:
                                sb.Append(cell.NumericCellValue.ToString() + "\t");
                                break;
                            case CellType.String:
                                sb.Append(cell.StringCellValue.Replace("\n", "") + "\t");
                                break;
                            default:
                                throw new Exception("?");
                        }
                    }
                    sb.Append("\r\n");
                }
            }
            return sb.ToString();
        }

        public void Write(TestTable table, string outFile)
        {
            var book = new XSSFWorkbook();
            var sheet = book.CreateSheet("data");

            //BorderStyle
            var borderStyle = book.CreateCellStyle();
            borderStyle.BorderTop = BorderStyle.Thin;
            borderStyle.BorderLeft = BorderStyle.Thin;
            borderStyle.BorderBottom = BorderStyle.Thin;
            borderStyle.BorderRight = BorderStyle.Thin;
            //style for DATE-Cell
            var dateStyle = book.CreateCellStyle();
            dateStyle.BorderTop = BorderStyle.Thin;
            dateStyle.BorderLeft = BorderStyle.Thin;
            dateStyle.BorderBottom = BorderStyle.Thin;
            dateStyle.BorderRight = BorderStyle.Thin;
            dateStyle.DataFormat = book.GetCreationHelper().CreateDataFormat().GetFormat("yyyy/mm/dd");

            for (int r = 0; r < table.RowCount; r++)
            {
                var datarow = table.GetNextRow();
                var destRow = sheet.CreateRow(r);
                ICell cell;
                var cellType = CellType.Blank;
                var colNum = 0;
                //STRING-Column
                //col1
                cellType = CellType.String;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col1);
                cell.CellStyle = borderStyle;
                colNum++;
                //col2
                cellType = CellType.String;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col2);
                cell.CellStyle = borderStyle;
                colNum++;
                //col3
                cellType = CellType.String;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col3);
                cell.CellStyle = borderStyle;
                colNum++;
                //col4
                cellType = CellType.String;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col4);
                cell.CellStyle = borderStyle;
                colNum++;
                //col5
                cellType = CellType.String;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col5);
                cell.CellStyle = borderStyle;
                colNum++;
                //NUMBER-Column
                //col6
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col6);
                cell.CellStyle = borderStyle;
                colNum++;
                //col7
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col7);
                cell.CellStyle = borderStyle;
                colNum++;
                //col8
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col8);
                cell.CellStyle = borderStyle;
                colNum++;
                //col9
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col9);
                cell.CellStyle = borderStyle;
                colNum++;
                //col10
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col10);
                cell.CellStyle = borderStyle;
                colNum++;
                //DATE-Column
                //col11
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col11);
                cell.CellStyle = dateStyle;
                colNum++;
                //col12
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col12);
                cell.CellStyle = dateStyle;
                colNum++;
                //col13
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col13);
                cell.CellStyle = dateStyle;
                colNum++;
                //col14
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col14);
                cell.CellStyle = dateStyle;
                colNum++;
                //col15
                cellType = CellType.Numeric;
                cell = destRow.CreateCell(colNum);
                cell.SetCellType(cellType);
                cell.SetCellValue(datarow.Col15);
                cell.CellStyle = dateStyle;
                colNum++;
            }

            //AutoFit
            for (var i = 0; i < table.ColCount; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            using (FileStream streamw = File.Open(outFile, FileMode.Create))
            {
                book.Write(streamw);
            }
        }
    }
}
