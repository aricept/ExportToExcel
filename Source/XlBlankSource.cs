using OfficeOpenXml;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace ExportToExcel
{
    public class XlBlankSource : IXlSource
    {
        private byte[] _data { get; }

        /// <summary>
        /// Generates a blank template when no template is provided.
        /// </summary>
        /// <param name="data">List of sheets to generate a report.</param>
        public XlBlankSource(IEnumerable<XlSheet> data)
        {
            using (var xl = new ExcelPackage())
            {
                // Loop through the data list and create a sheet for each one.
                foreach (var sheet in data)
                {
                    // Name each sheet if a name was provided, or use the model name.
                    sheet.Name ??= sheet.Type.Name;

                    // If a sheet with this name doesn't exist, create it. 
                    var worksheet = xl.Workbook.Worksheets[sheet.Name] ?? CreateSheet(xl, sheet);
                    AddRows(worksheet, sheet);
                    worksheet.Cells.AutoFitColumns();
                }

                _data = xl.GetAsByteArray();
            }
        }

        /// <summary>
        /// Load data from source.
        /// </summary>
        /// <returns><c>MemoryStream</c> of report data.</returns>
        public Stream Load()
        {
            return new MemoryStream(_data);
        }

        /// <summary>
        /// Required by <c>IXlSource</c> interface. Since this source always returns data, this method is always true.
        /// </summary>
        /// <returns></returns>
        public bool IsValid()
        {
            return true;
        }

        /// <summary>
        /// Creates a new sheet with headers based on objects in data list.
        /// </summary>
        /// <param name="xl">ExcelPackage for creating worksheet</param>
        /// <param name="sheet">The XlSheet object containing the sheet data.</param>
        /// <returns>The newly created worksheet</returns>
        public static ExcelWorksheet CreateSheet(ExcelPackage xl, XlSheet sheet)
        {
            // Create worksheet, set header style
            var worksheet = xl.Workbook.Worksheets.Add(sheet.Name);
            var headerStyle = worksheet.Row(1).Style;
            headerStyle.Font.Bold = true;
            headerStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            var data = sheet.Data().ToList();

            // Gets object properties based on type.
            var type = data[0].GetType();
            var props = type.GetProperties().Where(p => !p.XlIgnore(type)).ToList();

            // Loops through properties and adds property name or Display Name as column header.
            for (var col = 1; col <= props.Count; col++)
            {
                var prop = props[col - 1];

                var displayName = prop.GetDisplayName(type);
                if (!string.IsNullOrEmpty(displayName))
                {
                    worksheet.Cells[1, col].Value = displayName;
                }
                else
                {
                    worksheet.Cells[1, col].Value = props[col - 1].Name;
                }
            }

            return worksheet;
        }

        /// <summary>
        /// Adds rows of data to sheet
        /// </summary>
        /// <param name="worksheet">The worksheet object to add rows to.</param>
        /// <param name="data">The data to add to the sheet.</param>
        /// <returns>The resulting worksheet.</returns>
        private ExcelWorksheet AddRows(ExcelWorksheet worksheet, XlSheet data)
        {
            var sheetList = data.Data().ToList();
            var baseType = sheetList[0].GetType();
            var props = baseType.GetProperties().Where(p => !p.XlIgnore(baseType)).ToList();

            var firstBlank = worksheet.Dimension.Rows + 1;
            var row = firstBlank;
            for (var i = 0; i < sheetList.Count; i++)
            {
                for (var col = 1; col <= props.Count; col++)
                {
                    worksheet.Cells[row, col].Value = props[col - 1].GetValue(sheetList[row - firstBlank]);
                }
                row++;
            }

            var formatCol = 1;
            foreach (var prop in props)
            {
                var dataType = prop.GetDataType(baseType);
                if (dataType == DataType.Date || dataType == DataType.DateTime)
                {
                    worksheet.Cells[2, formatCol, sheetList.Count + 1, formatCol].Style.Numberformat.Format = "mm/dd/yyyy hh:mm";
                }
                formatCol++;
            }

            return worksheet;
        }
    }
}