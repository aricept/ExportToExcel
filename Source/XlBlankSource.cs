using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

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
                    sheet.Name = sheet.Name ?? sheet.Type.Name;
                    ExcelWorksheet worksheet;

                    // If a sheet with this name doesn't exist, create it. 
                    worksheet = xl.Workbook.Worksheets[sheet.Name] ?? CreateSheet(xl, sheet);
                    AddRows(worksheet, sheet);
                    worksheet.Cells.AutoFitColumns();
                }

                this._data = xl.GetAsByteArray();
            }
        }

        /// <summary>
        /// Load data from source.
        /// </summary>
        /// <returns><c>MemoryStream</c> of report data.</returns>
        public Stream Load()
        {
            return new MemoryStream(this._data);
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
        /// <param name="sheetName">The sheet name, provided by user or determined by type.</param>
        /// <returns>The newly created worksheet</returns>
        public ExcelWorksheet CreateSheet(ExcelPackage xl, XlSheet sheet)
        {
            // Create worksheet, set header style
            var worksheet = xl.Workbook.Worksheets.Add(sheet.Name);
            var headerStyle = worksheet.Row(1).Style;
            headerStyle.Font.Bold = true;
            headerStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

            // Gets object properties based on type.
            var props = sheet.Data().ToList()[0].GetType().GetProperties();

            // Loops through properties and adds property name or Display Name as column header.
            for (var col = 1; col <= props.Length; col++)
            {
                var prop = props[col - 1];
                var displayName = prop.GetDisplayName();
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
        public ExcelWorksheet AddRows(ExcelWorksheet worksheet, XlSheet data)
        {
            var sheetList = data.Data().ToList();
            var baseType = sheetList[0].GetType();
            var props = baseType.GetProperties();

            var firstBlank = worksheet.Dimension.Rows + 1;
            var row = firstBlank;
            for (var i = 0; i < sheetList.Count; i++)
            {
                for (var col = 1; col <= props.Length; col++)
                {
                    worksheet.Cells[row, col].Value = props[col - 1].GetValue(sheetList[row - firstBlank]);
                }
                row++;
            }

            return worksheet;
        }
    }
}