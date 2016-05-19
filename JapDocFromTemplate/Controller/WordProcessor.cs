using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Controller
{
    internal class WordProcessor
    {
        private readonly ThisDocument _doc;

        public WordProcessor(ThisDocument doc)
        {
            _doc = doc;
        }

        public int Start(string jsonSource)
        {
            var database = TableUtility.JsonToDocument(jsonSource);

            var tableCountDiff = database.Tables.Count - _doc.Tables.Count/2;
            while (tableCountDiff >= 0)
            {
                AddBlankPage();

                AddNewTable(_doc.Tables[2]);
                AddNewTable(_doc.Tables[1]);

                tableCountDiff = database.Tables.Count - _doc.Tables.Count/2;
            }

            TableUtility.GenerateTableDataJson(database, _doc.Tables);
            return _doc.Tables.Count;
        }

        private void AddBlankPage()
        {
            _doc.Words.Last.InsertBreak();
        }

        private void AddNewTable(Table table)
        {
            table.Range.Copy();
            _doc.Words.Last.Paste();
        }
    }
}