using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using SuperSimple.Spreadsheets.Serializer;

namespace SuperSimple.Spreadsheets
{
    public class ExcelSaver
    {
        private const string DEF_WORKSHEET_NAME = "Worksheet 1";
        #region Fields and properties

        IEnumerable<ExcelRow> DataToSave { get; set; }
        string SheetName { get; set; }
        #endregion

        public ExcelSaver(IEnumerable<ExcelRow> data, string sheetName = DEF_WORKSHEET_NAME)
        {
            DataToSave = data;
            SheetName = sheetName;
        }

        public void Save(Func<Stream> getStream, Func<Exception, bool> handleOnError = null, Action onSucces = null,
            bool disposeStreamAfterWrite = true)
        {
            Stream streamToReadFrom = null;

            try
            {
                streamToReadFrom = getStream();

                SaveToStream(streamToReadFrom);

                if(onSucces != null)
                {
                    onSucces();
                }
            }
            catch(Exception ex)
            {
                if (handleOnError != null)
                {
                    if(!handleOnError(ex))
                    {
                        throw;
                    }
                }
            }
            finally
            {
                if(streamToReadFrom != null && disposeStreamAfterWrite)
                {
                    streamToReadFrom.Dispose();
                }
            }
        }

        //TODO: (Savvas): Not a good method, but will do as part of the current task (OLYMPUS-1652)
        // Should really support this as an open-source project.
        public static void Save<T>(IEnumerable<T> rows, Stream streamToSaveTo = null, string sheetName = DEF_WORKSHEET_NAME, ISerializerToExcelRow serializer = null)
        {
            if(serializer == null)
            {
                serializer = new SerializerToExcelRow();
            }

            var data = serializer.Serialize(rows);

            var saver = new ExcelSaver(data, sheetName);

            saver.SaveToStream(streamToSaveTo);
        }

        public void SaveToStream(Stream stream)
        {
            //We create a temporary stream to support for operations 
            // that would cause an exception with network streams

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                var id = document.WorkbookPart.GetIdOfPart(worksheetPart);
                Sheet results = new Sheet()
                {
                    Id = id,
                    SheetId = 1,
                    Name = SheetName
                };
                sheets.Append(results);

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                foreach (var entry in DataToSave)
                    WriteRow(sheetData, entry);
            }

            stream.Flush();
        }

        #region Helper methods

        private void WriteRow(SheetData data, ExcelRow erow)
        {
            Cell old = null;

            Row row = new Row();
            data.Append(row);

            foreach (var cd in erow)
                AppendToRow(cd, row, ref old);
        }

        private void AppendToRow(ExcelCell data, Row row, ref Cell old)
        {
            Cell newCell = new Cell();
            row.InsertAfter(newCell, old);

            if (data.Value == null)
            {
                newCell.CellValue = new CellValue("");
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                old = newCell;
                return;
            }

            try
            {
                newCell.CellValue = new CellValue(data.Value.ToString(System.Globalization.CultureInfo.InvariantCulture.NumberFormat));
            }
            catch
            {
                newCell.CellValue = new CellValue("");
            }

            if (data.ValueType == typeof(DateTime))
            {
                newCell.CellValue = ToExcelDate(data.Value);
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            else if (data.ValueType == typeof(string))
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            else if (data.ValueType == typeof(int) || data.ValueType == typeof(long) || data.ValueType == typeof(float) || data.ValueType == typeof(double) || data.ValueType == typeof(decimal))
            {
                newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
            }
            else if (data.GetType() == typeof(DateTime))
            {
                newCell.DataType = new EnumValue<CellValues>(CellValues.Date);
            }

            old = newCell;
        }

        private CellValue ToExcelDate(object time)
        {
            Contract.Assert(time is DateTime);

            var from1900 = (DateTime)time - new DateTime(1900, 1, 1);

            return new CellValue(from1900.TotalDays.ToString());
        }
        #endregion

    }
}
