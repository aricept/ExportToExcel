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
        public XlFileInfo File { get; set; }

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
            _data = sheets.ToList();
            File = new XlFileInfo(name, sheets, method);
            using (var stream = File.FileSource.Load())
            {
                xl = new ExcelPackage(stream);
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
            _data = sheets.ToList();
            File = fileInfo;

            if (File.FileSource == null || !File.FileSource.IsValid())
            {
                File.FileSource = new XlBlankSource(sheets);
            }

            using (var stream = File.FileSource.Load())
            {
                xl = new ExcelPackage(stream);
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
            File = new XlFileInfo(newData.Type.Name, _data);
            using (var stream = File.FileSource.Load())
            {
                xl = new ExcelPackage(stream);
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
                if (File.FileSource.GetType().Equals(typeof(XlBlankSource)))
                {
                    break;
                }

                var baseType = report.Data().First().GetType();
                // ??= introduced in C# 8 assigns the right hand value to the left hand only if the left hand is null.
                // Here, if report.Name is null, assigns the baseType as the report name.
                report.Name ??= baseType.Name;
                ExcelWorksheet sheet;
                int row;

                if (xl.Workbook.Worksheets[report.Name] == null)
                {
                    sheet = XlBlankSource.CreateSheet(xl, report);//xl.Workbook.Worksheets.Add(sheetName);
                    row = 2;
                }
                else
                {
                    sheet = xl.Workbook.Worksheets[report.Name];
                    row = sheet.Dimension.End.Row + 1;
                }

                
                var dataList = report.Data().ToList();
                var props = baseType.GetProperties().Where(p => !p.XlIgnore(baseType)).ToList();

                for (var i = 0; i < dataList.Count; i++)
                {
                    for (var col = 1; col <= props.Count; col++)
                    {
                        sheet.Cells[row, col].Value = props[col - 1].GetValue(dataList[i]);
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

            return xlData;// File.Output.Save(xlData, File);
        }

        public void AddSheet(IEnumerable<object> data)
        {
            var newData = new XlSheet(data);
            _data.Add(newData);
        }
    }
}
