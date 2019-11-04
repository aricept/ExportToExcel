namespace ExportToExcel
{
    public interface IXlOutput
    {
        byte[] Save(byte[] data, XlFileInfo file = null);
    }
}