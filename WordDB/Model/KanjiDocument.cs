using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordDB.Model
{
    public class KanjiDocument
    {
        public int Count => Tables.Count;
        public List<Table> Tables { get; } = new List<Table>();
    }
}
