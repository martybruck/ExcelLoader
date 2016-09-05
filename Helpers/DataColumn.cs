using System.Collections.Generic;

namespace ExcelLoader.Helpers
{
    class DataColumn
    {
        public DataColumn( string colName)
        {
            ColumnName = colName;
        }

        public string ColumnName { get; set; }
        public int NumEntries
        {
            get
            {
                return (ColData == null ? 0 : ColData.Count);
            }
        }
        public List<string> ColData { get; set; } = new List<string>();

        public void Add(string colData)
        {
            ColData.Add(colData);
        }
    }
}
