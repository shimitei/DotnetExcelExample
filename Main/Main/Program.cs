using System;
using System.IO;

namespace Main
{
    class Program
    {
        /// <summary>
        /// Excel read-test:
        ///     Main.exe read ExcelDataReader
        ///     Main.exe read npoi
        ///     Main.exe read epplus
        /// Excel write-test:
        ///     Main.exe write npoi
        ///     Main.exe write epplus
        /// specified file(input)
        ///     Main.exe read npoi in:C:\test\test.xlsx
        /// specified dir(output)
        ///     Main.exe read npoi out:C:\test\
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            string inFile = CommandLineUtil.GetCommandOption("in:", @"C:\test\KEN_ALL.xlsx");
            string outFile = CommandLineUtil.GetCommandOption("out:", @"C:\test\");
            outFile += Path.DirectorySeparatorChar + "{0}{1}";
            if (CommandLineUtil.HasCommandOption("Read"))
            {
                ExecRead(inFile, outFile);
            } else if (CommandLineUtil.HasCommandOption("Write"))
            {
                ExecWrite(outFile);
            }
            PrintPeekMemory();
        }

        static void ExecRead(string inFile, string outFile)
        {
            var reader = GetExcelReaderWriter();
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            var data = reader.Read(inFile);
            sw.Stop();
            Console.WriteLine("{0} READ {1} msec", reader.GetName(), sw.ElapsedMilliseconds.ToString());

            //output read data
            File.WriteAllText(string.Format(outFile, reader.GetName(), ".txt"), data);
        }

        static void ExecWrite(string outFile)
        {
            var reader = GetExcelReaderWriter();
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            int row = int.Parse(CommandLineUtil.GetCommandOption("row:", "10000"));
            reader.Write(new TestTable(row), string.Format(outFile, reader.GetName(), ".xlsx"));
            sw.Stop();
            Console.WriteLine("{0} WRITE {1} msec", reader.GetName(), sw.ElapsedMilliseconds.ToString());
        }

        static IExcelReaderWriter GetExcelReaderWriter()
        {
            IExcelReaderWriter result = null;
            if (CommandLineUtil.HasCommandOption(ExcelDataReader.Id))
            {
                result = new ExcelDataReader();
            }
            else if (CommandLineUtil.HasCommandOption(NPOI.Id))
            {
                result = new NPOI();
            }
            else if (CommandLineUtil.HasCommandOption(EPPlus.Id))
            {
                result = new EPPlus();
            }
            return result;
        }

        static void PrintPeekMemory()
        {
            System.Diagnostics.Process currentProcess = System.Diagnostics.Process.GetCurrentProcess();
            currentProcess.Refresh();
            Console.WriteLine("PeakWorkingSet: {0}", currentProcess.PeakWorkingSet64);
            Console.WriteLine("PeakVirtualMemorySize: {0}", currentProcess.PeakVirtualMemorySize64);
        }
    }
}
