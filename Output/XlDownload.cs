namespace ExportToExcel
{
    public class XlDownload : IXlOutput
    {
        public byte[] Save(byte[] data, XlFileInfo file = null)
        {
            return data;
        }
    }
}
