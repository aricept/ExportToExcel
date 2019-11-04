namespace ExportToExcel
{
    public interface IXlOutput<T>
    {
        byte[] Save(byte[] data, XlFileInfo<T> file = null);
    }
}