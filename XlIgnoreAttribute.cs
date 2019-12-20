using System;

namespace ExportToExcel
{
    [AttributeUsageAttribute(AttributeTargets.Property)]
    public class XlIgnoreAttribute : Attribute
    {
        public XlIgnoreAttribute() { }
    }
}