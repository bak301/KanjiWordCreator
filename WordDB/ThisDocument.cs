using System;
using WordDB.Controller;
using Office = Microsoft.Office.Core;

namespace WordDB
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, EventArgs e)
        {
            const string path = @"C:\Users\vn130\OneDrive\Documents\Word Document\dist\Tap3.docx";
            var wordProcessor = new WordProcessor(this);
            var document = Application.Documents.Open(path);
            var tables = document.Tables;

            //wordProcessor.Start(tables, 7);
            //wordProcessor.WriteJsonToTextFile(tables, 7, "boThu.json");
            //wordProcessor.WriteJsonToTextFile(tables, 5, "300.json");
            wordProcessor.WriteJsonToTextFile(tables, 5, "kanji_part2.json");
            document.Close();
        }

        private void ThisDocument_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisDocument_Startup;
            Shutdown += ThisDocument_Shutdown;
        }

        #endregion
    }
}