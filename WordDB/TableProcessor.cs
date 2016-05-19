using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace WordDB
{
    internal class TableProcessor
    {
        private readonly Document _doc;
        private int _rowCount;

        public TableProcessor(Document doc, int rowCount)
        {
            _doc = doc;
            _rowCount = rowCount;
        }

        public IEnumerable<string> ReadAllTable()
        {
            var result = new List<string>();
            Debug.WriteLine($"Tables count : {_doc.Tables.Count}");
            for (var i = 1; i <= _doc.Tables.Count; i += 2)
            {
                var kanjiTable = _doc.Tables[i];
                var hanVietTable = _doc.Tables[i + 1];

                if (kanjiTable.Rows.Count < _rowCount && hanVietTable.Rows.Count < _rowCount)
                    _rowCount = kanjiTable.Rows.Count;

                for (var j = 1; j <= _rowCount; j++)
                {
                    var kanjiRow = kanjiTable.Rows[j];
                    var hanVietRow = hanVietTable.Rows[j];

                    var kanjiData = ReadKanji(kanjiRow);
                    var hanVietData = ReadHanViet(hanVietRow).ToList();
                    hanVietData = hanVietData.Select(e => e.Replace(" ", "-")).ToList();

                    var lineData = MergeLine(kanjiData, hanVietData);
                    result.Add(lineData);
                }

                _rowCount = 5;
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
    }
}