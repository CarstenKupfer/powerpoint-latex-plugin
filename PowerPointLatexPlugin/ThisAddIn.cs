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
        public bool WaitingForMathField { get; set; } = false;

        // NEW: handles our own LaTeX field drawing mode
        private LatexFieldDrawer latexFieldDrawer;


        private void Application_WindowSelectionChange(
            Microsoft.Office.Interop.PowerPoint.Selection selection)
        {
            if (!WaitingForMathField)
                return;

            if (selection.Type != Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
                return;

            if (selection.ShapeRange.Count != 1)
                return;

            Microsoft.Office.Interop.PowerPoint.Shape shape =
                selection.ShapeRange[1];

            if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                shape.Tags.Add("MathField", "true");

                System.Windows.Forms.MessageBox.Show(
                    "Math Field successfully detected!\nTag: "
                    + shape.Tags["MathField"]
                );

                WaitingForMathField = false;
            }
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange +=
                Application_WindowSelectionChange;

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