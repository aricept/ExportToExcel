namespace ExportToExcel
{
    public class XlDownload<T> : IXlOutput<T>
    {
        public byte[] Save(byte[] data, XlFileInfo<T> file = null)
        {
            return data;
        }
    }
}
