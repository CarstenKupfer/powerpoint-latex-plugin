using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointLatexPlugin
{
    public partial class ThisAddIn
    {
        // NEW: handles our own LaTeX field drawing mode
        private LatexFieldDrawer latexFieldDrawer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // NEW
            latexFieldDrawer =
                new LatexFieldDrawer(this.Application);
        }


        // NEW: called by our Insert LaTeX button
        public void StartLatexFieldDrawing()
        {
            latexFieldDrawer.Start();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // NEW
            latexFieldDrawer?.Dispose();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}