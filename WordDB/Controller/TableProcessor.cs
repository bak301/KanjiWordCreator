using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using JapDocFromTemplate.Model;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using Row = Microsoft.Office.Interop.Word.Row;

namespace WordDB.Controller
{
    internal class TableProcessor
    {
        public string TablesToJson(Tables tables, int rowCount)
        {
            Debug.WriteLine($"Tables count : {tables.Count}");

            var kanjiDocument = new Model.KanjiDocument();
            for (var i = 1; i < tables.Count; i += 2)
            {
                var kanjiTable = tables[i];
                var hanVietTable = tables[i + 1];

                if (kanjiTable.Rows.Count < rowCount && hanVietTable.Rows.Count < rowCount)
                    rowCount = kanjiTable.Rows.Count;

                var kanjiCharacterTable = new Model.Table();
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

        private Model.Row ParseRow(Row kanjiRow, Row hanVietRow)
        {
            //char[] charToRemove = {'\r', '\u0007'};
            var result = from Cell kanjiCell in kanjiRow.Cells
                         join Cell hanVietCell in hanVietRow.Cells 
                         on kanjiCell.ColumnIndex equals hanVietCell.ColumnIndex
                         let hanVietWord = hanVietCell.Range.Paragraphs.First.Range.Text
                         select new KanjiCharacter
                         {
                             Kanji = kanjiCell.Range.Text[0],
                             HanViet = Regex.Replace(hanVietWord, "[\r\u0007]", ""),
                             Meaning = Regex.Replace(hanVietCell.Range.Text, $@"[\r\u0007]|({hanVietWord})","")
                         };

            return new Model.Row
            {
                KanjiCharacters = result
            };
        }

        #region Old Method
        public IEnumerable<string> ReadAllTable(Tables tables, int rowCount)
        {
            var result = new List<string>();
            Debug.WriteLine($"Tables count : {tables.Count}");
            for (var i = 1; i <= tables.Count; i += 2)
            {
                var kanjiTable = tables[i];
                var hanVietTable = tables[i + 1];

                if (kanjiTable.Rows.Count < rowCount && hanVietTable.Rows.Count < rowCount)
                    rowCount = kanjiTable.Rows.Count;

                for (var j = 1; j <= rowCount; j++)
                {
                    var kanjiRow = kanjiTable.Rows[j];
                    var hanVietRow = hanVietTable.Rows[j];

                    var kanjiData = ReadKanji(kanjiRow);
                    var hanVietData = ReadHanViet(hanVietRow).ToList();
                    hanVietData = hanVietData.Select(e => e.Replace(" ", "-")).ToList();

                    var lineData = MergeLine(kanjiData, hanVietData);
                    result.Add(lineData);
                }

                rowCount = 5;
            }
            return result;
        }

        public IEnumerable<string> ReadKanji(Row r)
        {
            var lineData = (from Cell cell in r.Cells select cell.Range.Text.Trim()).ToList();
            lineData.RemoveAll(string.IsNullOrWhiteSpace);
            return lineData;
        }

        public IEnumerable<string> ReadHanViet(Row r)
        {
            var lineData = from Cell cell in r.Cells select cell.Range;
            var result =
                (from range in lineData
                    let firstWord = range.Paragraphs.First.Range.Text
                    select range.Text.Insert(firstWord.Length, " ")).ToList();
            return result;
        }

        public Dictionary<string, string> GetRowDictionary(List<string> kanji, List<string> hanViet)
        {
            if (kanji.Count != hanViet.Count)
                throw new Exception($"Two list don't have the same number of elements :" +
                                    $"\nHanViet Count = {hanViet.Count} " +
                                    $"\nKanji Count = {kanji.Count}");

            return kanji.Zip(hanViet, (k, v) => new {Key = k, Value = v})
                .ToDictionary(x => x.Key.ToString(), x => x.Value);
        }

        public string MergeLine(IEnumerable<string> kanji, IEnumerable<string> hanViet)
        {
            var stringBuilder = new StringBuilder();

            foreach (var s in kanji)
            {
                stringBuilder.Append(s);
            }

            foreach (var s in hanViet)
            {
                stringBuilder.Append($" {s}");
            }

            var result = stringBuilder.ToString().Replace("\r", "");
            return result;
        }
#endregion


    }
}