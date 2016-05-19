using System;
using System.Diagnostics;
using JapDocFromTemplate.Controller;

namespace JapDocFromTemplate
{
    public partial class ThisDocument
    {
        private void Process_Source(string fileName)
        {
            var processor = new WordProcessor(this);
            var tablesPopulated = processor.Start(fileName);
            Debug.WriteLine($"Total number of tables : {tablesPopulated}");
        }

        private void ThisDocument_Startup(object sender, EventArgs e)
        {
            Process_Source("300.json");
            Exporter.WriteJsonToTextFile(this.Tables, 5, "recentlySaved-300.json");
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