using System;
using System.Collections.Generic;
using System.Linq;


namespace ExportToExcel
{
    /// <summary>
    /// Class for transmitting information about data into the exporter.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class XlSheet<T>
    {
        private List<T> _data;
        public string Name { get; set; }

        /// <summary>
        /// The current model type. Used to set sheet name when none is provided.
        /// </summary>
        public Type Type
        {
            get
            {
                return typeof(T);
            }
        }

        /// <summary>
        /// Creates empty XlSheet. Name can be set by using <c>Name</c>, and data can be added using <c>Data()</c> method.
        /// </summary>
        public XlSheet()
        {
            _data = new List<T>();
        }

        /// <summary>
        /// Contains data the exporter uses to create report.
        /// </summary>
        /// <param name="name">The Worksheet tab that the data will be added to</param>
        /// <param name="data">List of strongly typed objects from which dat will be pulled to fill the report.</param>
        public XlSheet(string name, IEnumerable<T> data)
        {
            Name = name;
            _data = data.ToList();
        }

        /// <summary>
        /// Contains data the exporter uses to create report. Without a provided name, the name will later be determined using the model name.
        /// </summary>
        /// <param name="data">List of strongly typed objects from which dat will be pulled to fill the report.</param>
        public XlSheet(IEnumerable<T> data)
        {
            _data = data.ToList();
        }

        /// <summary>
        /// Retrieve current data.
        /// </summary>
        /// <returns>The current list of data.</returns>
        public IEnumerable<T> Data()
        {
            return _data;
        }

        /// <summary>
        /// Add a new range of data.
        /// </summary>
        /// <param name="newData">List of new data to be added.</param>
        public void Data(IEnumerable<T> newData)
        {
            _data.AddRange(newData);
        }
    }
}
