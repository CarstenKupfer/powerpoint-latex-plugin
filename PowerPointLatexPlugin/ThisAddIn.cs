using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLatexPlugin
{
    public partial class ThisAddIn
    {

        // NEW: handles our own LaTeX field drawing mode
        private LatexFieldDrawer latexFieldDrawer;

        private PowerPoint.Shape activeLatexField = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // NEW
            latexFieldDrawer =
                new LatexFieldDrawer(this.Application);

            this.Application.WindowSelectionChange +=
                Application_WindowSelectionChange;
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

        private void Application_WindowSelectionChange(
            PowerPoint.Selection selection)
        {
            PowerPoint.Shape currentLatexField = null;

            try
            {
                if (selection.Type ==
                        PowerPoint.PpSelectionType.ppSelectionShapes ||
                    selection.Type ==
                        PowerPoint.PpSelectionType.ppSelectionText)
                {
                    if (selection.ShapeRange.Count == 1)
                    {
                        PowerPoint.Shape shape =
                            selection.ShapeRange[1];

                        if (shape.Tags["LatexField"] == "true")
                        {
                            currentLatexField = shape;
                        }
                    }
                }
            }
            catch
            {
                // Aktives Feld nicht verlieren
            }


            // 1. Altes LaTeX-Feld wurde verlassen → Source speichern
            if (activeLatexField != null)
            {
                bool stillSameField =
                    currentLatexField != null &&
                    currentLatexField.Id == activeLatexField.Id;

                if (!stillSameField)
                {
                    try
                    {
                        string latexSource =
                            activeLatexField.TextFrame.TextRange.Text;

                        activeLatexField.Tags.Add(
                            "LatexSource",
                            latexSource
                        );

                        System.Windows.Forms.MessageBox.Show(
                            "LaTeX source saved!\n\n" +
                            "LatexSource = " + latexSource
                        );
                    }
                    catch
                    {
                    }

                    activeLatexField = null;
                }
            }


            // 2. Neues LaTeX-Feld wurde ausgewählt
            if (currentLatexField != null)
            {
                bool isNewSelection =
                    activeLatexField == null ||
                    activeLatexField.Id != currentLatexField.Id;

                if (isNewSelection)
                {
                    activeLatexField = currentLatexField;

                    try
                    {
                        string savedLatex =
                            currentLatexField.Tags["LatexSource"];

                        // Nur anzeigen, wenn bereits etwas gespeichert wurde
                        if (!string.IsNullOrEmpty(savedLatex))
                        {
                            System.Windows.Forms.MessageBox.Show(
                                "Loaded LaTeX source:\n\n" +
                                savedLatex
                            );
                        }
                    }
                    catch
                    {
                    }
                }
            }
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