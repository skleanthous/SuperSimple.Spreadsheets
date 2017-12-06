using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SuperSimple.Spreadsheets.ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var spreadsheet = new ExcelLoader("test.xlsx", false))
            {
                spreadsheet.Data
                    .Select(x => x.Aggregate($"CellCount:{x.Count} -> ", (o, c) => String.Format("{0}|{1}", o, c?.Value ?? "")))
                    .ToList()
                    .ForEach(Console.WriteLine);
            }

            Console.ReadKey();
        }

    }
}
