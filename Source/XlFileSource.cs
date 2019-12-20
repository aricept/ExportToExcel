using System.Configuration;
using System.IO;
using System.Web.Hosting;

namespace ExportToExcel
{
    public class XlFileSource : IXlSource
    {
        private string _path { get; set; }
        private bool isValid { get; }

        /// <summary>
        /// Provides source to a template file on disk.
        /// </summary>
        /// <param name="path">The path to the template file. May be an AppSettings key, virtual path, or absolute path.</param>
        public XlFileSource(string path)
        {
            _path = GetFilePathIfExists(path, out bool isValid);
            this.isValid = isValid;
        }

        /// <summary>
        /// Loads template from file.
        /// </summary>
        /// <returns></returns>
        public Stream Load()
        {
            return File.Open(_path, FileMode.Open, FileAccess.ReadWrite);
        }

        /// <summary>
        /// Private method used to parse the source directory.
        /// </summary>
        /// <param name="source">Provided by client, may be an AppSettings key, virtual path, or absolute path.</param>
        /// <param name="fileExists">Output parameter determining if file was found.</param>
        /// <returns></returns>
        private string GetFilePathIfExists(string source, out bool fileExists)
        {
            fileExists = false;
            var truePath = string.Empty;
            var tempPath = ConfigurationManager.AppSettings[source];

            if (tempPath == null)
            {
                tempPath = source;
            }

            tempPath = HostingEnvironment.MapPath(tempPath);

            if (File.Exists($"{tempPath}"))
            {
                truePath = tempPath;
                fileExists = true;
            }
            else if (File.Exists(source))
            {
                truePath = source;
                fileExists = true;
            }

            return truePath;
        }

        /// <summary>
        /// Is the XlFileSource valid?
        /// </summary>
        /// <remarks>This method is used to determine if XlExporter can actually load the provided source or if it should create a new XlBlankSource instead.
        /// This way we do not throw a <c>FileNotFoundException</c>, but at least return a report.</remarks>
        /// <returns>True if valid, False if not.</returns>
        public bool IsValid()
        {
            return isValid;
        }
    }
}
