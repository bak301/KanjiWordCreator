using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace WordDB.Controller
{
    internal class WordProcessor
    {
        private readonly ThisDocument _currentDoc;
        private readonly TableProcessor _tableProcessor;

        public WordProcessor(ThisDocument currentDoc)
        {
            _currentDoc = currentDoc;
            _tableProcessor = new TableProcessor();
        }

        public void Start(Tables tables, int rowCount)
        {
            var allLines = _tableProcessor.ReadAllTable(tables, rowCount).ToList();
            for (var i = 1; i <= allLines.Count; i++)
            {
                WriteNewLine(allLines[i - 1]);
                if (i%rowCount != 0) continue;
                _currentDoc.Paragraphs.Last.Range.InsertParagraphAfter();
            }
            _currentDoc.Paragraphs.First.Range.Delete();
        }

        public void WriteJsonToTextFile(Tables tables, int rowCount, string name)
        {
            var data = _tableProcessor.TablesToJson(tables, rowCount);
            File.WriteAllText($@"C:\Users\vn130\OneDrive\Documents\Word Document\src\Database\{name}", data);
        }

        public void WriteNewLine(string data)
        {
            _currentDoc.Paragraphs.Last.Range.InsertParagraphAfter();
            _currentDoc.Paragraphs.Last.Range.Text = data;
        }
    }
}