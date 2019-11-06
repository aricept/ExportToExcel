using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace ExportToExcel
{
    public class XlExporter
    {
        private ExcelPackage xl { get; set; }
        private byte[] xlData { get; set; }
        private List<XlSheet> _data { get; set; }
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
            this._data = sheets.ToList();
            this._file = new XlFileInfo(name, sheets, method);
            using (var stream = this._file.FileSource.Load())
            {
                this.xl = new ExcelPackage(stream);
            }
        }

        /// <summary>
        /// Exports data to an Excel spreadsheet.
        /// </summary>
        /// <param name="sheets">List of individual sheets to be added.</param>
        /// <param name="fileInfo">The XlFileInfo object containing information about the source and destination of the data.</param>
        /// <param name="selected">Optional Tuple containing the sheetname and cell address to be selected when opening the file.</param>
        public XlExporter(IEnumerable<XlSheet> sheets, XlFileInfo fileInfo, (string sheet, string cell)? selected = null)
        {
            this._data = sheets.ToList();
            this._file = fileInfo;

            if (_file.FileSource == null || !_file.FileSource.IsValid())
            {
                _file.FileSource = new XlBlankSource(sheets);
            }

            using (var stream = this._file.FileSource.Load())
            {
                this.xl = new ExcelPackage(stream);
            }

            if (selected != null)
            {
                OpenSelect = selected;
            }
        }

        public XlExporter(IEnumerable<object> data)
        {
            var newData = new XlSheet(data);
            _data = new List<XlSheet> { newData };
            _file = new XlFileInfo(newData.Type.Name, _data);
            using (var stream = this._file.FileSource.Load())
            {
                this.xl = new ExcelPackage(stream);
            }
        }

        /// <summary>
        /// Runs the report, including running the XlFileInfo's <c>Save</c> method.
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
                int row;

                if (xl.Workbook.Worksheets[report.Name] == null)
                {
                    sheet = xl.Workbook.Worksheets.Add(report.Name);
                    row = 1;
                }
                else
                {
                    sheet = xl.Workbook.Worksheets[report.Name];
                    row = sheet.Dimension.End.Row + 1;
                }

                var baseType = report.Data().First().GetType();
                var dataList = report.Data().ToList();
                var props = baseType.GetProperties();

                for (var i = 0; i < dataList.Count; i++)
                {
                    for (var col = 1; col <= props.Length; col++)
                    {
                        sheet.Cells[row, col].Value = props[col-1].GetValue(dataList[i]);
                    }
                    row++;
                }

                sheet.Cells.AutoFitColumns();
            }

            if (OpenSelect != null)
            {
                xl.Workbook.Worksheets[OpenSelect.Value.sheet].Select(OpenSelect.Value.cell);
            }

            xlData = xl.GetAsByteArray();

            return _file.Output.Save(xlData, _file);
        }

        public void AddSheet(IEnumerable<object> data)
        {
            var newData = new XlSheet(data);
            _data.Add(newData);
        }
    }
}
