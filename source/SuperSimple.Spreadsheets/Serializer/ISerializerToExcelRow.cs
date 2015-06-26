using System;
using System.Collections.Generic;

namespace SSExcel.Serializer
{
    public interface ISerializerToExcelRow
    {
        IEnumerable<ExcelRow> Serialize<T>(IEnumerable<T> itemsToSerialize, bool getHeaders = true);
    }
}
