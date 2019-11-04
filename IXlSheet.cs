using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportToExcel
{
    public interface IXlSheet
    {
        IEnumerable Data();
        string Name { get; set; }
        string Type { get; set; }
    }
}
