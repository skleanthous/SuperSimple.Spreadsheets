using System;
using System.Collections.Generic;

namespace SuperSimple.Spreadsheets.Serializer
{
    public interface ISerializerToExcelRow
    {
        IEnumerable<ExcelRow> Serialize<T>(IEnumerable<T> itemsToSerialize, bool getHeaders = true);
    }
}
