using System;
using System.Linq;

namespace Main
{
    public class TestTable
    {
        public int RowIndex { get; set; }
        public int RowCount { get; private set; }
        public int ColCount { get; private set; }

        public TestRow GetNextRow()
        {
            RowIndex++;
            return new TestRow(RowIndex);
        }
        public TestTable(int rowCount = 10000)
        {
            RowCount = rowCount;
            ColCount = 15;
            RowIndex = 0;
        }
    }

    public class TestRow
    {
        private int index;
        private Random rnd;

        public TestRow(int index)
        {
            this.index = index;
            rnd = new Random(index);
        }

        public int RowNum { get { return index; } }
        public object[] GetArray()
        {
            var result = new object[] {
                    Col1,
                    Col2,
                    Col3,
                    Col4,
                    Col5,
                    Col6,
                    Col7,
                    Col8,
                    Col9,
                    Col10,
                    Col11,
                    Col12,
                    Col13,
                    Col14,
                    Col15,
                };
            return result;
        }
        public string Col1 { get { return String1; } }
        public string Col2 { get { return String2; } }
        public string Col3 { get { return String3; } }
        public string Col4 { get { return String4; } }
        public string Col5 { get { return String5; } }
        public int Col6 { get { return Number1; } }
        public int Col7 { get { return Number2; } }
        public int Col8 { get { return Number3; } }
        public int Col9 { get { return Number4; } }
        public int Col10 { get { return Number5; } }
        public DateTime Col11 { get { return Date1; } }
        public DateTime Col12 { get { return Date2; } }
        public DateTime Col13 { get { return Date3; } }
        public DateTime Col14 { get { return Date4; } }
        public DateTime Col15 { get { return Date5; } }

        private string String1 { get { return GetString(index); } }
        private string String2 { get { return GetString(index); } }
        private string String3 { get { return GetString(index); } }
        private string String4 { get { return GetString(index); } }
        private string String5 { get { return GetString(index); } }
        private int Number1 { get { return GetNumber(index); } }
        private int Number2 { get { return GetNumber(index); } }
        private int Number3 { get { return GetNumber(index); } }
        private int Number4 { get { return GetNumber(index); } }
        private int Number5 { get { return GetNumber(index); } }
        private DateTime Date1 { get { return GetDate(index); } }
        private DateTime Date2 { get { return GetDate(index); } }
        private DateTime Date3 { get { return GetDate(index); } }
        private DateTime Date4 { get { return GetDate(index); } }
        private DateTime Date5 { get { return GetDate(index); } }

        private int GetNumber(int num)
        {
            return rnd.Next(1, 123456789);
        }
        private char[] strdef = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz".ToArray();
        private string GetString(int num)
        {
            var result = "";
            const int maxlen1 = 20;
            for (var i = 0; i < maxlen1; i++)
            {
                result += strdef[rnd.Next(1, strdef.Length) - 1];
            }
            return result.Substring(rnd.Next(0, result.Length - 4));
        }
        private DateTime GetDate(int num)
        {
            return new DateTime(1990 + rnd.Next(0, 30), rnd.Next(1, 12), rnd.Next(1, 28));
        }
    }
}
