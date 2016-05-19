using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Controller
{
    internal static class Exporter
    {
        public static void WriteJsonToTextFile(Tables tables, int rowCount, string name)
        {
            var data = TableUtility.TablesToJson(tables, rowCount);
            File.WriteAllText($@"C:\Users\vn130\OneDrive\Documents\Word Document\src\Database\Json\{name}", data);
        }
    }
}
