using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Controller
{
    internal class WordProcessor
    {
        private readonly ThisDocument _doc;
        private readonly TableProcessor _tableProcessor;

        public WordProcessor(ThisDocument doc, string fileName, int rowCount)
        {
            _doc = doc;
            _tableProcessor = new TableProcessor(_doc.Application, fileName, rowCount);
        }

        public int GenerateData(int numberOfPages)
        {
            var sourceRowIndexStart = 0;
            for (var i = 0; i < numberOfPages; i++)
            {
                AddBlankPage();

                AddNewTable(_doc.Tables[2]);
                AddNewTable(_doc.Tables[1]);

                var index = _doc.Tables.Count;
                _tableProcessor.GenerateTableData(ref sourceRowIndexStart, _doc.Tables[index], _doc.Tables[index - 1]);
            }
            return _doc.Tables.Count;
        }

        private void AddBlankPage()
        {
            _doc.Words.Last.InsertBreak();
        }

        private void AddNewTable(Table table)
        {
            table.Range.Copy();
            //_doc.Paragraphs.Last.Range.Paste();
            _doc.Words.Last.Paste();
        }

        // Experimental method
    }
}