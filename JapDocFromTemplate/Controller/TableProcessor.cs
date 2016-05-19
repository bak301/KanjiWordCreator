using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Controller
{
    internal class TableProcessor
    {
        private readonly StringProcessor _dataProcessor;

        public TableProcessor(Application app, string fileName)
        {
            _dataProcessor = new StringProcessor(app, fileName);
        }
        
        //Experimental
        //public int GenerateTableDataTuple(int sourceRowIndex, int rowCount, Table hanVietTable, Table kanjiTable)
        //{
        //    for (var row = 1; row <= rowCount; row++)
        //    {
        //        var lineData = _dataProcessor.GetKanjiTuple(sourceRowIndex + row);
        //        var enumerable = lineData as Tuple<char, string, string>[] ?? lineData.ToArray();

        //        var kanjiList = enumerable.Select(e => e.Item1.ToString()).ToList();
        //        var hanVietList = enumerable.Select(e => e.Item2).ToList();
        //        var meaningList = enumerable.Select(e => e.Item3).ToList();

        //        GenerateRowData(kanjiList, kanjiTable, row);
        //        GenerateRowData(hanVietList, hanVietTable, row);
        //        GenerateRowData(meaningList, hanVietTable, row);
        //    }

        //    return sourceRowIndex + rowCount + 1;
        //}

        public int GenerateTableData(ref int kanjiCount, int sourceRowIndex, int rowCount, Table hanVietTable, Table kanjiTable)
        {
            for (var row = 1; row <= rowCount; row++)
            {
                var lineData = _dataProcessor.GetKanjiDictionary(sourceRowIndex + row);
                var kanjiList = lineData.Keys;
                kanjiCount += kanjiList.Count;
                var hanVietList = lineData.Values;

                GenerateRowData(kanjiList, kanjiTable, row, false);
                GenerateRowData(hanVietList, hanVietTable, row, true);
            }

            return sourceRowIndex + rowCount + 1;
        }

        private void GenerateRowData(ICollection<string> collection, Table table, int rowIndex,bool isHanViet)
        {
            if (collection.Count > table.Columns.Count)
            {
                AddMoreColumns(table, collection.Count);
            }

            foreach (Column col in table.Columns)
            {
                var index = col.Index;

                try
                {
                    var currentCell = col.Cells[rowIndex];
                    var data = collection.ElementAt(index - 1);

                    GenerateCellData(currentCell, data, isHanViet);
                }
                catch (ArgumentOutOfRangeException)
                {
                }
            }
        }

        private void GenerateCellData(Cell cell, string data, bool isHanViet)
        {
            cell.Range.Text += data;
            cell.Range.Text = cell.Range.Text.Replace("\r", "");

            //if (isHanViet)
            //{
            //    var wordCount = cell.Range.Words.Count;
            //}
        }


        private void AddMoreColumns(Table table, int count)
        {
            var difference = count - table.Columns.Count;

            for (var i = 0; i < difference; i++)
            {
                table.Columns.Add();
                RecalibrateTableSize(table, count);
            }
        }

        private void RecalibrateTableSize(Table table, int count)
        {
            var percentage = (table.Columns.Count - 1)/(float) table.Columns.Count;

            if (count == 7)
            {
                table.Range.Font.Size *= percentage;
                table.Range.Font.Size += (float) 0.5;
            }

            for (var i = 1; i <= table.Columns.Count; i++)
            {
                table.Columns[i].Width *= percentage;
            }
        }
    }
}