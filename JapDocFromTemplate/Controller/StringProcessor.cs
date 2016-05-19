using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using JapDocFromTemplate.Model;
using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Controller
{
    internal class StringProcessor
    {
        private readonly DocumentRepo _docRepo;

        public StringProcessor(Application app, string fileName)
        {
            _docRepo = new DocumentRepo(app, $@"{fileName}");
        }

        // ---------------- Get Dictionary
        public Dictionary<string, string> GetKanjiDictionary(int paragraphIndex)
        {
            var lineDataArray = LineToArray(paragraphIndex);

            var kanjiList = lineDataArray[0].ToList();
            var hanVietList = lineDataArray
                .Skip(1)
                .Select(str => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(str))
                .ToList();

            var result = kanjiList
                .Zip(hanVietList, (k, v) => new {Key = k, Value = v})
                .ToDictionary(x => x.Key.ToString(), x => x.Value);

            return result;
        }

        private string[] LineToArray(int index)
        {
            var lineData = GetLineByParagraph(index);
            char[] delimiter = {' ', (char) 160};
            var result = lineData
                .Split(delimiter, StringSplitOptions.RemoveEmptyEntries);

            return result;
        }

        private string GetLineByParagraph(int index)
        {
            var line = _docRepo.Source.Paragraphs[index].Range.Text;
            line = line.Remove(line.Length - 1);
            //Debug.WriteLine($"Line data : {line}");
            return line;
        }

        #region EXPERIMENTAL METHOD

        public IEnumerable<KanjiCharacter> GetKanji(int paragraphIndex)
        {
            var result = new List<KanjiCharacter>();
            var lineDataArray = LineToArray(paragraphIndex);

            var kanjiList = lineDataArray[0].ToList();
            var hanVietList = lineDataArray
                .Skip(1)
                .Select(str => str.Split('-')[0]);
            var meaningList = lineDataArray
                .Skip(1)
                .Select(str => string.Join(" ", str.Split('-').Skip(1)));

            var e1 = kanjiList.GetEnumerator();
            var e2 = hanVietList.GetEnumerator();
            var e3 = meaningList.GetEnumerator();

            while (e1.MoveNext() && e2.MoveNext() && e3.MoveNext())
            {
                var kanji = e1.Current;
                var hanViet = e2.Current;
                var mean = e3.Current;

                result.Add(new KanjiCharacter
                {
                    Kanji = kanji,
                    HanViet = hanViet,
                    Meaning = mean
                });
            }
            return result;
        }

        #endregion
    }
}