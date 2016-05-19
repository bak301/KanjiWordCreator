using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JapDocFromTemplate.Model;

namespace WordDB.Model
{
    public class Row
    {
        public int Count => KanjiCharacters.Count();
        public IEnumerable<KanjiCharacter> KanjiCharacters { get; set; }
    }
}
