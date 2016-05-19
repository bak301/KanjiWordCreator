using System.IO;
using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Model
{
    internal class Repository
    {
        public Repository(string jsonPath)
        {
            JsonSource =
                File.ReadAllText($@"C:\Users\vn130\OneDrive\Documents\Word Document\src\Database\{jsonPath}");
        }

        public string JsonSource { get; set; }
    }
}