using Main;
using System;
using System.IO;

namespace OpenXmlPowerToolsCustom
{
    class Program
    {
        /// <summary>
        /// Excel write-test:
        ///     OpenXmlPowerToolsCustom.exe write OpenXmlPowerTools
        /// specified dir(output)
        ///     OpenXmlPowerToolsCustom.exe write OpenXmlPowerTools out:C:\test\
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            string inFile = CommandLineUtil.GetCommandOption("in:", @"C:\test\KEN_ALL.xlsx");
            string outFile = CommandLineUtil.GetCommandOption("out:", @"C:\test\");
            outFile += Path.DirectorySeparatorChar + "{0}{1}";
            if (CommandLineUtil.HasCommandOption("Write"))
            {
                ExecWrite(outFile);
            }
            PrintPeekMemory();
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
            if (CommandLineUtil.HasCommandOption(OXPT.Id))
            {
                result = new OXPT();
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
