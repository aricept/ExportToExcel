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
                    var sheetName = sheet.Name ?? sheet.Type.Name;
                    ExcelWorksheet worksheet;
                    int firstBlank;
                    bool newSheet = false;

                    // If a sheet with this name doesn't exist, create it. 
                    if (xl.Workbook.Worksheets[sheetName] == null)
                    {
                        worksheet = xl.Workbook.Worksheets.Add(sheetName);
                        firstBlank = 1;
                        newSheet = true;
                    }
                    // If it exists, find the first blank row.
                    else
                    {
                        worksheet = xl.Workbook.Worksheets[sheetName];
                        firstBlank = worksheet.Dimension.Rows + 1;
                    }

                    var sheetList = sheet.Data().ToList();
                    var baseType = sheetList[0].GetType();
                    var props = baseType.GetProperties();

                    // If this is a new sheet, set the header styles to make text centered and bold and add headers.
                    if (newSheet)
                    {
                        var headerStyle = worksheet.Row(firstBlank).Style;
                        headerStyle.Font.Bold = true;
                        headerStyle.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        
                        for (var col = 1; col <= props.Length; col++)
                        {
                            var prop = props[col - 1];
                            var displayName = prop.GetDisplayName();
                            if (!string.IsNullOrEmpty(displayName))
                            {
                                worksheet.Cells[firstBlank, col].Value = displayName;
                            }
                            else
                            {
                                worksheet.Cells[firstBlank, col].Value = props[col - 1].Name;
                            }
                        }
                        firstBlank++;
                    }

                    // Load the data into the sheet and autofit columns to the data.
                    var row = firstBlank;
                    for (var i = 0; i < sheetList.Count; i++)
                    {
                        for (var col = 1; col <= props.Length; col++)
                        {
                            worksheet.Cells[row, col].Value = props[col - 1].GetValue(sheetList[row - firstBlank]);
                        }
                        row++;
                    }

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
    }
}