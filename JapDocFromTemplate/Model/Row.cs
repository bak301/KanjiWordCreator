using System.Collections.Generic;
using System.Linq;

namespace JapDocFromTemplate.Model
{
    public class Row
    {
        public int Count => KanjiCharacters.Count();
        public IEnumerable<KanjiCharacter> KanjiCharacters { get; set; }
    }
}