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

        public Dictionary<string, string> GetKanjiDictionary(int paragraphIndex)
        {
            var lineDataArray = LineToArray(paragraphIndex);

            var kanjiList = lineDataArray[0].ToList();
            var hanVietList = lineDataArray
                .Skip(1)
                .Select(str => CultureInfo.CurrentCulture.TextInfo.ToTitleCase(str))
                .ToList();

            if (hanVietList.Count != kanjiList.Count)
                throw new Exception($"Two list don't have the same number of elements :" +
                                    $"\nHanViet Count = {hanVietList.Count} " +
                                    $"\nKanji Count = {kanjiList.Count}");

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
            return line;
        }
    }
}