using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSExcel
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

        public void AddCell(dynamic value)
        {
            this.Add(new ExcelCell(value));
        }
    }
}
