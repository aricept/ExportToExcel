using System.IO;

namespace ExportToExcel
{
    public interface IXlSource
    {
        Stream Load();
        bool IsValid();
    }
}
