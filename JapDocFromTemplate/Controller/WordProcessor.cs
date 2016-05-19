using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;

namespace JapDocFromTemplate.Controller
{
    internal class WordProcessor
    {
        private readonly ThisDocument _doc;

        public WordProcessor(ThisDocument doc)
        {
            _doc = doc;
        }

        public int Start(string fileName)
        {
            var database = TableUtility.JsonToDocument(fileName);

            var tableCountDiff = database.Tables.Count - _doc.Tables.Count / 2;
            while (tableCountDiff >= 0)
            {
                AddBlankPage();

                AddNewTable(_doc.Tables[2]);
                AddNewTable(_doc.Tables[1]);
                Debug.WriteLine($"Number of tables : {_doc.Tables.Count}");

                tableCountDiff = database.Tables.Count - _doc.Tables.Count / 2;
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
            //_doc.Paragraphs.Last.Range.Paste();
            _doc.Words.Last.Paste();
        }
    }
}