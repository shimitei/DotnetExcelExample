using System;
using System.Data;
using System.IO;
using System.Text;

namespace Main
{
    class ExcelDataReader : IExcelReaderWriter
    {
        public const string Id = "ExcelDataReader";
        public string GetName()
        {
            return Id;
        }

        public string Read(string inFile)
        {
            var stream = File.OpenRead(inFile);
            var excelReader = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
            var dataSet = excelReader.AsDataSet();
            stream.Close();

            var sb = new StringBuilder();
            foreach (DataRow dataRow in dataSet.Tables[0].Rows)
            {
                foreach (var obj in dataRow.ItemArray)
                {
                    sb.Append(obj.ToString().Replace("\n", "") + "\t");
                }
                sb.Append("\r\n");
            }
            excelReader.Close();
            return sb.ToString();
        }

        public void Write(TestTable table, string outFile)
        {
            // no support
            throw new NotImplementedException();
        }
    }
}
