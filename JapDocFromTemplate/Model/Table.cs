using System.Collections.Generic;

namespace JapDocFromTemplate.Model
{
    public class Table
    {
        public int Count => Rows.Count;
        public List<Row> Rows { get; } = new List<Row>();
    }
}