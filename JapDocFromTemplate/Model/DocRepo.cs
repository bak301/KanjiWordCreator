using Microsoft.Office.Interop.Word;

namespace JapDocFromTemplate.Model
{
    internal class DocumentRepo
    {
        public DocumentRepo(_Application app, string filePath)
        {
            Source = app.Documents.Open($@"C:\Users\vn130\OneDrive\Documents\Word Document\{filePath}");
        }

        public Document Source { get; private set; }
    }
}