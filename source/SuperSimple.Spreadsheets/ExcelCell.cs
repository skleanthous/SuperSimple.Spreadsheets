using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;

namespace SuperSimple.Spreadsheets
{
    public class ExcelCell
    {
        public Type ValueType
        { get; set; }
        public dynamic Value
        { get; set; }

        public ExcelCell(dynamic value)
        {
            if (value == null)
            {
                Value = "";
                ValueType = typeof(string);
            }
            else
            {
                Value = value;
                ValueType = value.GetType();
            }
        }
    }
}
