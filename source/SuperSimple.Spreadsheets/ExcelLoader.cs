using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace SuperSimple.Spreadsheets
{
    public class ExcelLoader : IDisposable
    {
        #region Fields and properties
        /// <summary>
        /// The separator character that will be used for processing. This should not exist in the data to be parsed.
        /// </summary>
        public const char SEPARATOR = '|';

        /// <summary>
        /// A flag indicating that ExcelLoader is disposed.
        /// </summary>
        public bool IsDisposed { get; private set; }

        /// <summary>
        /// The excel spreadsheet document.
        /// </summary>
        SpreadsheetDocument Document { get; set; }

        public bool IgnoreNullOrEmptyCells { get; }

        bool changedSheetProcessingFunc = true;
        private Func<Sheet, bool> confirmSheetProcessing;
        /// <summary>
        /// A func to confirm each that each sheet should be processed.
        /// </summary>
        public Func<Sheet, bool> ConfirmSheetProcessing
        {
            get { return confirmSheetProcessing; }
            set
            {
                if (value != confirmSheetProcessing)
                {
                    confirmSheetProcessing = value;
                    changedSheetProcessingFunc = true;
                }
            }
        }

        private List<ExcelRow> rowStringData = null;
        /// <summary>
        /// Returns the rows in the excel file as a list of rows, where each row is a list of strings.
        /// </summary>
        public List<ExcelRow> Data
        {
            get
            {
                return changedSheetProcessingFunc ? rowStringData = ReadRows() : rowStringData;
            }
        }

        /// <summary>
        /// The file that was opened.
        /// </summary>
        public String Filename { get; private set; }

        #region Workbook and parts iterators and accessors
        /// <summary>
        /// The main iteratetor to retrieve each workbook part which contains sheets. Automatically sets the shared string table, active sheets and stylesheet.
        /// </summary>
        public IEnumerable<WorkbookPart> WorkbookParts
        {
            get
            {
                foreach (var workbookPart in Document.GetPartsOfType<WorkbookPart>())
                {
                    SharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    Sheets = new List<Sheet>();

                    foreach (var s in workbookPart.Workbook.Sheets.Where(x => x is Sheet))
                    {
                        var sheet = s as Sheet;
                        if (sheet != null)
                            ((List<Sheet>)Sheets).Add(sheet);
                    }

                    if(workbookPart.WorkbookStylesPart != null)
                    {
                        Stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
                    }

                    ActiveWorkbookPart = workbookPart;

                    yield return workbookPart;
                }
            }
        }

        /// <summary>
        /// The main iterator to retrieve all SheetData from the document.
        /// </summary>
        public IEnumerable<SheetData> SheetData
        {
            get
            {
                foreach (var book in WorkbookParts)
                    foreach (var data in ActiveSheetData)
                        yield return data;
            }
        }

        /// <summary>
        /// The main iterator to retrieve all SheetData from the document.
        /// </summary>
        public IEnumerable<Tuple<SheetData,SharedStringTablePart, Stylesheet>> SheetDataWithMeta
        {
            get
            {
                foreach (var book in WorkbookParts)
                    foreach (var data in ActiveSheetData)
                        yield return new Tuple<SheetData,SharedStringTablePart, Stylesheet>(data, SharedStringTable, Stylesheet);
            }
        }

        /// <summary>
        /// The main iterator to retrieve all SheetData from the active WorkbookPart retrieved from the WorkbookPart iterator.
        /// </summary>
        private IEnumerable<SheetData> ActiveSheetData
        {
            get
            {
                foreach (Sheet sheet in Sheets)
                    foreach (SheetData sheetdata in (ActiveWorkbookPart.GetPartById(((Sheet)sheet).Id) as WorksheetPart).Worksheet.Where(x => x is SheetData))
                    {
                        if (ConfirmSheetProcessing != null && !ConfirmSheetProcessing(sheet))
                            continue;

                        yield return sheetdata;
                    }
            }
        }

        /// <summary>
        /// The active WorkbookPart from the WorkbookPart iterator.
        /// </summary>
        private WorkbookPart ActiveWorkbookPart
        { get; set; }

        /// <summary>
        /// The shared string table of the active WorkbookPart retrieved from the WorkbookParts iterator.
        /// </summary>
        private SharedStringTablePart SharedStringTable
        { get; set; }

        /// <summary>
        /// The sheets contained in the active WorkbookPart retrieved from the WorkbookParts iterator.
        /// </summary>
        private IEnumerable<Sheet> Sheets
        { get; set; }

        /// <summary>
        /// The stylesheet of the active WorkbookPart from the WorkbookParts iterator.
        /// </summary>
        private Stylesheet Stylesheet
        { get; set; }
        #endregion

        #endregion

        #region Ctors

        /// <summary>
        /// Loads an excel 2007 file for reading only.
        /// </summary>
        /// <param name="path"></param>
        public ExcelLoader(string path, bool ignoreNullOrEmptyCells = true)
            : this(SpreadsheetDocument.Open(path, false), ignoreNullOrEmptyCells)
        {
            Filename = path.Contains("\\") ? path.Substring(path.LastIndexOf('\\') + 1) : path;
        }

        /// <summary>
        /// An excel 2007 file for reading, already open in readonly mode.
        /// </summary>
        /// <param name="document"></param>
        public ExcelLoader(SpreadsheetDocument document, bool ignoreNullOrEmptyCells = true)
        {
            Document = document;
            IgnoreNullOrEmptyCells = ignoreNullOrEmptyCells;
        }
        #endregion

        #region Row processing

        public List<ExcelRow> ReadRows()
        {
            List<ExcelRow> tableData = new List<ExcelRow>();

            foreach (var dataWithMeta in SheetDataWithMeta)
                foreach(Row row in dataWithMeta.Item1.ChildElements)
                {
                    //List<string> rowData = new List<string>();
                    ExcelRow rowData = new ExcelRow();
                    foreach (Cell c in row.ChildElements)
                    {
                        if (c == null || c.CellValue == null || string.IsNullOrWhiteSpace(c.CellValue.Text))
                        {
                            if(!IgnoreNullOrEmptyCells) rowData.AddCell(null);

                            continue;
                        }

                        var styles = dataWithMeta.Item3;

                        //Dates from excel: http://blogs.msdn.com/b/eric_carter/archive/2004/08/14/214713.aspx
                        ExcelCell val = null;

                        CellFormat toFindNumbFormat = null;
                        if(c.StyleIndex != null && styles != null)
                            toFindNumbFormat = styles.CellFormats.ToArray()[c.StyleIndex.Value] as CellFormat;

                        bool isDate = false;

                        if (toFindNumbFormat != null && toFindNumbFormat.NumberFormatId.HasValue && toFindNumbFormat.NumberFormatId.Value != 0)
                        {
                            var index = toFindNumbFormat.NumberFormatId.Value;
                            if (index >= 163)
                            {
                                NumberingFormat format = styles.NumberingFormats.First(x => x is NumberingFormat && ((NumberingFormat)x).NumberFormatId == toFindNumbFormat.NumberFormatId.Value) as NumberingFormat;

                                if (format != null && format.FormatCode.HasValue && VerifyDateFormatCode(format.FormatCode.Value))
                                    isDate = true;
                            }
                            else if ((index >= 14 && index <= 22) || (index >= 45 && index <= 47))
                                isDate = true;
                        }

                        double doubleVal;
                        long intVal;
                        if (c.CellValue == null)
                            continue;
                        if (c.DataType != null && c.DataType == CellValues.SharedString)
                            val = new ExcelCell(dataWithMeta.Item2.SharedStringTable.ElementAt(int.Parse(c.CellValue.Text)).InnerText);
                        else if (isDate)
                            val = new ExcelCell(GetDateFromExcelDate(c.CellValue.Text));
                        else if (c.DataType != null && c.DataType == CellValues.Date)
                            val = new ExcelCell(DateTime.Parse(c.CellValue.Text));
                        else if (long.TryParse(c.CellValue.Text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out intVal))
                            val = new ExcelCell(intVal);
                        else if (double.TryParse(c.CellValue.Text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out doubleVal))
                            val = new ExcelCell(doubleVal);
                        else
                            val = new ExcelCell(c.CellValue.Text);


                        rowData.Add(val);
                    }

                    tableData.Add(rowData);
                }

            return tableData;
        }

        #endregion

        #region Verify data and helpers

        /// <summary>
        /// Verifies that the given numbering format string represents a date --> crude for now, but efficient and effective.
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private bool VerifyDateFormatCode(string p)
        {
            return p.Contains('m') || p.Contains('y') || p.Contains('d');
        }

        /// <summary>
        /// Retrieves a date frin the string data that represents a date. This only works for dates AFTER 1904.
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        private DateTime GetDateFromExcelDate(string p)
        {
            return new DateTime(1900, 1, 1) + TimeSpan.FromDays(double.Parse(p, System.Globalization.NumberFormatInfo.InvariantInfo)) - TimeSpan.FromDays(2);
        }
        #endregion

        public static ExcelLoader LoadReadOnlyFromStream(Stream stream)
        {
            return new ExcelLoader(SpreadsheetDocument.Open(stream, true));
        }

        #region IDispose pattern
        ~ExcelLoader()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        public void Dispose(bool disposing)
        {
            if (!IsDisposed)
            {
                if (Document != null)
                    Document.Dispose();
            }
        }
        #endregion
    }
}
