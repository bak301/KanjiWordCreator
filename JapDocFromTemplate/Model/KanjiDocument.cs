using System.Collections.Generic;

namespace JapDocFromTemplate.Model
{
    public class KanjiDocument
    {
        public int Count => Tables.Count;
        public List<Table> Tables { get; } = new List<Table>();
    }
}