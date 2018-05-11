using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;

namespace SuperSimple.Spreadsheets
{
    public class ExcelRow: List<ExcelCell>
    {
        public ExcelRow()
        {
        }

        public ExcelRow(params object[] values)
        {
            foreach (var val in values)
                this.Add(new ExcelCell(val));
        }

        /// <summary>
        /// This is to match iterators
        /// </summary>
        /// <param name="values"></param>
        public ExcelRow(IEnumerable<object> values)
        {
            foreach (var val in values)
                this.Add(new ExcelCell(val));
        }

        public void AddCell(dynamic value)
        {
            this.Add(new ExcelCell(value));
        }
    }
}
