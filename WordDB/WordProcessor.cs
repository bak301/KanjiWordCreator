using System.Linq;
using Microsoft.Office.Interop.Word;

namespace WordDB
{
    internal class WordProcessor
    {
        private readonly ThisDocument _currentDoc;
        private readonly int _rowCount;
        private readonly TableProcessor _tableProcessor;

        public WordProcessor(ThisDocument currentDoc, Document doc, int rowCount)
        {
            _currentDoc = currentDoc;
            _rowCount = rowCount;
            _tableProcessor = new TableProcessor(doc, _rowCount);
        }

        public void Start()
        {
            var allLines = _tableProcessor.ReadAllTable().ToList();
            for (var i = 1; i <= allLines.Count; i++)
            {
                WriteNewLine(allLines[i - 1]);
                if (i%_rowCount != 0) continue;
                _currentDoc.Paragraphs.Last.Range.InsertParagraphAfter();
            }
            _currentDoc.Paragraphs.First.Range.Delete();
        }

        public void WriteNewLine(string data)
        {
            _currentDoc.Paragraphs.Last.Range.InsertParagraphAfter();
            _currentDoc.Paragraphs.Last.Range.Text = data;
        }
    }
}