using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportToExcel
{
    /// <summary>
    /// Contains data about the base file, the save file, and backup file.
    /// </summary>
    public class XlFileInfo<T>
    {
        /// <summary>
        /// The template path. If one is not provided a blank template will be created.
        /// </summary>
        public IXlSource FileSource { get; set; }
        public string FileName { get; set; }
        public string BackupPath { get; set; }
        public IXlOutput<T> Output { get; set; }
        private ExcelPackage xl;

        public XlFileInfo() { }

        /// <summary>
        /// Contains data relating to the file-transfer aspects of the report. This overload is used when there is no template file being used, and no backup directory is provided.
        /// </summary>
        /// <param name="name">The FileName to use when saving the report.</param>
        /// <param name="data">List of sheets to create the report.</param>
        /// <param name="method">The optional save method: Local or Download.</param>
        public XlFileInfo(string name, IEnumerable<XlSheet<T>> data, XlSaveMethod method = XlSaveMethod.Local)
        {
            FileName = name;

            if (method == XlSaveMethod.Download)
            {
                Output = new XlDownload<T>();
            }
            else
            {
                Output = new XlDownloadAndBackup<T>();
            }

            FileSource = new XlBlankSource<T>(data);
        }

        /// <summary>
        /// Contains data relating to the file-transfer aspects of the report. This overload is used when there is no template file being used, and providing a backup diretcory.
        /// </summary>
        /// <param name="name">The FileName to use when saving ethe report.</param>
        /// <param name="directory">How to locate the save directory. This may be an AppSettings key, a virtual path, or the absolute path to the directory.</param>
        /// <param name="data">List of sheets to create the report.</param>
        public XlFileInfo(string name, string directory, IEnumerable<XlSheet<T>> data)
        {
            FileName = name;
            Output = new XlDownloadAndBackup<T>();
            FileSource = new XlBlankSource<T>(data);

            BackupPath = Utils.GetDirPathIfExists(directory, out bool exists);

            if (!exists)
            {
                throw new DirectoryNotFoundException(
                    $"No backup directory could be found using '{directory}' checking in AppSettings, " +
                    "Virtual Paths, and Absolute Paths. Please provide a valid key for AppSettings, " +
                    "a valid Virtual Path, or a valid Absolute Path.");
            }
        }

        /// <summary>
        /// Contains data relating to the file-transfer aspects of the report.
        /// </summary>
        /// <param name="source">The template filename. This may be an AppSettings key, a virtual path, or just the filename if it exists.</param>
        /// <param name="name">The filename to use when saving/downloading the report.</param>
        /// <param name="method">The optional save method to use: Local or Download.</param>
        public XlFileInfo(string source, string name, XlSaveMethod method = XlSaveMethod.Local)
        {
            FileName = name;

            if (method == XlSaveMethod.Download)
            {
                Output = new XlDownload<T>();
            }
            else
            {
                Output = new XlDownloadAndBackup<T>();
            }

            FileSource = new XlFileSource(source);
        }

        /// <summary>
        /// Used to download and backup report.
        /// </summary>
        /// <param name="source">The template filename. This may be an AppSettings key, a virtual path, or just the filename if it exists.</param>
        /// <param name="name">The filename to use when saving/downloading the report.</param>
        /// <param name="backup">The diretcory to save backup to. This may be an AppSettings key, a virtual directory, or a full path.</param>
        public XlFileInfo(string source, string name, string backup)
        {
            FileName = name;
            Output = new XlDownloadAndBackup<T>();
            FileSource = new XlFileSource(source);

            BackupPath = Utils.GetDirPathIfExists(backup, out var backupExists);

            if (!backupExists)
            {
                throw new DirectoryNotFoundException(
                    $"No backup directory could be found using '{backup}' checking in AppSettings, " +
                    "Virtual Paths, and Absolute Paths. Please provide a valid key for AppSettings, " +
                    "a valid Virtual Path, or a valid Absolute Path.");
            }

        }
    }
}
