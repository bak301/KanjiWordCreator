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
                var newestTables = AddPageWithTables();
                _tableProcessor.GenerateTableData(ref sourceRowIndexStart, newestTables[1], newestTables[0]);
            }
            return _doc.Tables.Count;
        }

        private Table[] AddPageWithTables()
        {
            AddBlankPage();

            var templateKanjiTable = _doc.Tables[1];
            var templateHanVietTable = _doc.Tables[2];

            var newTableHanTu = AddNewTable(templateKanjiTable);
            var newTableHanViet = AddNewTable(templateHanVietTable);
            var newest = new[] {newTableHanTu, newTableHanViet};

            return newest;
        }

        private void AddBlankPage()
        {
            _doc.Words.Last.InsertBreak();
        }

        private Table AddNewTable(Table table)
        {
            table.Range.Copy();
            _doc.Words.Last.Paste();
            var lastTableIndex = _doc.Tables.Count;
            return _doc.Tables[lastTableIndex];
        }
    }
}