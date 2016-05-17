using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Controller
{
    internal class TableProcessor
    {
        private readonly StringProcessor _dataProcessor;
        private readonly int _rowCount;

        public TableProcessor(Application app, string fileName, int rowCount)
        {
            _dataProcessor = new StringProcessor(app, fileName);
            _rowCount = rowCount;
        }

        public void GenerateTableData(ref int sourceRowIndex, Table tblHanViet, Table tblHanTu)
        {
            for (var row = 1; row <= _rowCount; row++)
            {
                var lineData = _dataProcessor.GetKanjiDictionary(sourceRowIndex + row);
                var kanjiList = lineData.Keys;
                var hanVietList = lineData.Values;

                GenerateRowData(kanjiList, tblHanTu, row);
                GenerateRowData(hanVietList, tblHanViet, row);
            }

            sourceRowIndex += _rowCount + 1;
        }

        private void GenerateRowData(ICollection<string> collection, Table table, int row)
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
                    GenerateCellData(col.Cells[row], collection.ElementAt(index - 1));
                }
                catch (ArgumentOutOfRangeException)
                {
                }
            }
        }

        private void GenerateCellData(Cell cell, string data)
        {
            cell.Range.Text = data;
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