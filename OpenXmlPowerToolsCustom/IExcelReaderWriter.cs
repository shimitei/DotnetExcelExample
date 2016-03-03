using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Main
{
    interface IExcelReaderWriter
    {
        string GetName();
        string Read(string inFile);
        void Write(TestTable table, string outFile);
    }
}
