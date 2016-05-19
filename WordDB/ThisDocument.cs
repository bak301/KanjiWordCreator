using System;
using Office = Microsoft.Office.Core;

namespace WordDB
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, EventArgs e)
        {
            const string path = @"C:\Users\vn130\OneDrive\Documents\Word Document\Pre_release\Table-bo-thu.docx";
            var wordProcessor = new WordProcessor(this, Application.Documents.Open(path), 5);
            wordProcessor.Start();
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