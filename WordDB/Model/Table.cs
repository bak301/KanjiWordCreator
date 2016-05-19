using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDB.Model
{
    public class Table
    {
        public int Count => Rows.Count;
        public List<Row> Rows { get; } = new List<Row>();
    }
}
