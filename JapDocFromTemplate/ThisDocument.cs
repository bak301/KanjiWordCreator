﻿using System;
using System.Diagnostics;
using JapDocFromTemplate.Controller;

namespace JapDocFromTemplate
{
    public partial class ThisDocument
    {
        private void Process_Source(string fileName, int rowCount, int pageCount)
        {
            var processor = new WordProcessor(this, fileName);
            var tablesPopulated = processor.GenerateData(pageCount, rowCount);

            Debug.WriteLine($"Number of tables : {tablesPopulated}");
        }

        private void ThisDocument_Startup(object sender, EventArgs e)
        {
            //Process_Source(@"Source\Source_v4.docx", 5, 51);
            //Process_Source(@"Source\BoThu.docx", 7, 5);
            //Process_Source(@"Source\first300.docx", 5, 12);
            //Process_Source("Source\Source_v4_File2.docx", 5, 38);

            Process_Source(@"Source\Database\first-300.source.docx", 5, 12);
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