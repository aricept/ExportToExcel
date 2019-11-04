using System.IO;

namespace ExportToExcel
{
    public class XlDownloadAndBackup<T> : IXlOutput<T>
    {
        /// <summary>
        /// Saves the report data to disk.
        /// </summary>
        /// <param name="data">The report data to be saved.</param>
        /// <param name="file">The XlFileINfo object containing the information about where to save the report.</param>
        /// <returns></returns>
        public byte[] Save(byte[] data, XlFileInfo<T> file)
        {
            File.WriteAllBytes($"{file.BackupPath}{file.FileName}", data);

            return data;
        }
    }
}
