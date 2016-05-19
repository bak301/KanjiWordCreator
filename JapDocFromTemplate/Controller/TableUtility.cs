using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using JapDocFromTemplate.Model;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Row = Microsoft.Office.Interop.Word.Row;
using Jap = JapDocFromTemplate.Model;

namespace JapDocFromTemplate.Controller
{
    internal static class TableUtility
    {
        //Experiment
        public static KanjiDocument JsonToDocument(string path)
        {
            var data = new Repository(path).JsonSource;
            var result = JsonConvert.DeserializeObject<KanjiDocument>(data);

            return result;
        }

        public static void GenerateTableDataJson(KanjiDocument db, Tables tables)
        {
            foreach (var table in db.Tables)
            {
                var tableIndex = db.Tables.IndexOf(table);
                Debug.WriteLine($"Current table index {tableIndex}");
                foreach (var row in table.Rows)
                {
                    var rowIndex = table.Rows.IndexOf(row);
                    Debug.WriteLine($" - Current row index {rowIndex}");
                    foreach (var data in row.KanjiCharacters)
                    {
                        var columnIndex = row.KanjiCharacters.ToList().IndexOf(data);
                        Debug.WriteLine($" - - Current column index {columnIndex}");

                        var tableKanji = tables[2*tableIndex + 3];
                        var tableHanViet = tables[2*tableIndex + 4];

                        var kanjiCell = tableKanji.Cell(rowIndex + 1, columnIndex + 1);
                        var hanVietCell = tableHanViet.Cell(rowIndex + 1, columnIndex + 1);

                        kanjiCell.Range.Text = data.Kanji.ToString();

                        var firstParagraph = hanVietCell.Range.Paragraphs.First ?? hanVietCell.Range.Paragraphs.Add();
                        Debug.WriteLine($"Bold value: {firstParagraph.Range.Font.Bold}");
                        firstParagraph.Range.Text = data.HanViet;
                        hanVietCell.Range.Text += data.Meaning;

                        firstParagraph.Range.Font.Size = 12;
                        firstParagraph.Range.Font.Bold = 0;
                    }
                }
            }
        }

        public static string TablesToJson(Tables tables, int rowCount)
        {
            Debug.WriteLine($"Tables count : {tables.Count}");

            var kanjiDocument = new Jap.KanjiDocument();
            for (var i = 1; i < tables.Count; i += 2)
            {
                var kanjiTable = tables[i];
                var hanVietTable = tables[i + 1];

                if (kanjiTable.Rows.Count < rowCount && hanVietTable.Rows.Count < rowCount)
                    rowCount = kanjiTable.Rows.Count;

                var kanjiCharacterTable = new Jap.Table();
                for (var j = 1; j <= rowCount; j++)
                {
                    var kanjiRow = kanjiTable.Rows[j];
                    var hanVietRow = hanVietTable.Rows[j];

                    var kanjiCharacterRow = ParseRow(kanjiRow, hanVietRow);
                    kanjiCharacterTable.Rows.Add(kanjiCharacterRow);
                }
                kanjiDocument.Tables.Add(kanjiCharacterTable);
            }
            return JsonConvert.SerializeObject(kanjiDocument, Formatting.Indented);
        }

        private static Jap.Row ParseRow(Row kanjiRow, Row hanVietRow)
        {
            //char[] charToRemove = {'\r', '\u0007'};
            var result = from Cell kanjiCell in kanjiRow.Cells
                         join Cell hanVietCell in hanVietRow.Cells
                             on kanjiCell.ColumnIndex equals hanVietCell.ColumnIndex
                         let hanVietWord = hanVietCell.Range.Paragraphs.First.Range.Text
                         let kanjiWord = kanjiCell.Range.Text[0]
                         where !kanjiWord.Equals('\r')
                         select new Jap.KanjiCharacter
                         {
                             Kanji = kanjiWord,
                             HanViet = Regex.Replace(hanVietWord, "[\r\u0007]", ""),
                             Meaning = Regex.Replace(hanVietCell.Range.Text, $@"[\r\u0007]|({hanVietWord})", "")
                         };

            return new Jap.Row
            {
                KanjiCharacters = result
            };
        }
    }
}