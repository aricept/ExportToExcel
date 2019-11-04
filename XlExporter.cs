using OfficeOpenXml;
using System.Collections.Generic;

namespace ExportToExcel
{
    public class XlExporter<T>
    {
        private ExcelPackage xl { get; set; }
        private byte[] xlData { get; set; }
        private IEnumerable<XlSheet> _data { get; set; }
        private XlFileInfo _file { get; set; }

        /// <summary>
        /// Tuple for providing optional selected cell on file open.
        /// </summary>
        public (string sheet, string cell)? OpenSelect { get; set; }
        public string Type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>
        /// Exports data to an Excel spreadsheet.
        /// </summary>
        /// <param name="sheets">List of individual sheets to be added.</param>
        /// <param name="name">Filename to be used when saving/downloading.</param>
        /// <param name="method">The optional save method: Local or Download.</param>
        public XlExporter(IEnumerable<XlSheet> sheets, string name, XlSaveMethod method = XlSaveMethod.Local)
        {
            this._data = sheets;
            this._file = new XlFileInfo(name, sheets, method);
            this.xl = new ExcelPackage(_file.FileSource.Load());
        }

        /// <summary>
        /// Exports data to an Excel spreadsheet.
        /// </summary>
        /// <param name="sheets">List of individual sheets to be added.</param>
        /// <param name="fileInfo">The XlFileInfo object containing information about the source and destination of the data.</param>
        /// <param name="selected">Optional Tuple containing the sheetname and cell address to be selected when opening the file.</param>
        public XlExporter(IEnumerable<XlSheet> sheets, XlFileInfo fileInfo, (string sheet, string cell)? selected = null)
        {
            this._data = sheets;
            this._file = fileInfo;

            if (_file.FileSource == null || !_file.FileSource.IsValid())
            {
                _file.FileSource = new XlBlankSource(sheets);
            }
            
            this.xl = new ExcelPackage(_file.FileSource.Load());

            if (selected != null)
            {
                OpenSelect = selected;
            }
        }

        /// <summary>
        /// Runs the report, including runing the XlFileInfo's <c>Save</c> method.
        /// </summary>
        /// <returns>A <c>byte[]</c> of the file data. This can be diverted elsewhere for saving in a different location or used to download.</returns>
        public byte[] Run()
        {
            foreach (var report in _data)
            {
                if (_file.FileSource.GetType().Equals(typeof(XlBlankSource)))
                {
                    break;
                }

                ExcelWorksheet sheet;
                int end;

                if (xl.Workbook.Worksheets[report.Name] == null)
                {
                    sheet = xl.Workbook.Worksheets.Add(report.Name);
                    end = 1;
                }
                else
                {
                    sheet = xl.Workbook.Worksheets[report.Name];
                    end = sheet.Dimension.End.Row + 1;
                }

                sheet.Cells["A" + end].LoadFromCollection(report.Data());
                sheet.Cells.AutoFitColumns();
            }

            if (OpenSelect != null)
            {
                xl.Workbook.Worksheets[OpenSelect.Value.sheet].Select(OpenSelect.Value.cell);
            }

            xlData = xl.GetAsByteArray();

            return _file.Output.Save(xlData, _file);
        }
    }
}
