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

        private bool isSwitchingLatexMode = false;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // NEW
            latexFieldDrawer =
                new LatexFieldDrawer(this.Application);

            latexFieldDrawer.DrawingFinished += () =>
            {
                Globals.Ribbons.MathRibbon.ResetInsertLatexToggle();
            };

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

            if (isSwitchingLatexMode)
                return;

            try
            {
                if ((selection.Type ==
                    PowerPoint.PpSelectionType.ppSelectionShapes || 
                    selection.Type ==
                    PowerPoint.PpSelectionType.ppSelectionText)
                    && selection.ShapeRange.Count == 1)
                {
                    PowerPoint.Shape selectedShape =
                        selection.ShapeRange[1];

                    if (selectedShape.Tags["LatexRender"] == "true")
                    {
                        OpenLatexEditor(selectedShape);
                        return;
                    }
                }
            }
            catch
            {
            }

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

                        string imagePath =
                               LatexRenderer.RenderToSvg(latexSource);      

                        PowerPoint.Slide slide = this.Application.ActiveWindow.View.Slide;

                        DeleteExistingRender(slide, activeLatexField.Id);

                        PowerPoint.Shape renderedFormula =
                            slide.Shapes.AddPicture(
                                imagePath,
                                Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoTrue,
                                activeLatexField.Left,
                                activeLatexField.Top,
                                -1,
                                -1
                            );

                        // Seitenverhältnis fixieren
                        renderedFormula.LockAspectRatio = Office.MsoTriState.msoTrue;

                        // Größe des ursprünglichen LaTeX-Feldes
                        float maxWidth = activeLatexField.Width;
                        float maxHeight = activeLatexField.Height;

                        // Aktuelle natürliche Bildgröße
                        float imageWidth = renderedFormula.Width;
                        float imageHeight = renderedFormula.Height;

                        // Skalierungsfaktoren berechnen
                        float widthScale = maxWidth / imageWidth;
                        float heightScale = maxHeight / imageHeight;

                        // Kleineren Faktoren nehmen, damit nichts abgeschnitten/verzerrt wird
                        float scaleFactor = Math.Min(widthScale, heightScale);

                        // Größe proportional setzen
                        renderedFormula.Width = imageWidth * scaleFactor;

                        // Durch LockAspectRatio passt sich die Höhe automatisch an

                        // Im ursprünglichen Feld zentrieren
                        renderedFormula.Left =
                            activeLatexField.Left +
                            (maxWidth - renderedFormula.Width) / 2;

                        renderedFormula.Top =
                            activeLatexField.Top +
                            (maxHeight - renderedFormula.Height) / 2;

                        // Rendering kennzeichnen
                        renderedFormula.Tags.Add("LatexRender", "true");

                        renderedFormula.Tags.Add(
                            "SourceShapeId",
                            activeLatexField.Id.ToString()
                        );

                        // LaTeX-Quelltext verstecken
                        activeLatexField.Visible = Office.MsoTriState.msoFalse;  
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
                    }
                    catch
                    {
                    }
                }
            }
        }

        private void OpenLatexEditor(PowerPoint.Shape renderedShape)
        {
            string sourceShapeIdText = renderedShape.Tags["SourceShapeId"];

            if (!int.TryParse(sourceShapeIdText, out int sourceShapeId))
                return;

            PowerPoint.Slide slide = this.Application.ActiveWindow.View.Slide;

            PowerPoint.Shape sourceShape = null;

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (shape.Id == sourceShapeId)
                {
                    sourceShape = shape;
                    break;
                }
            }

            if (sourceShape == null)
                return;

            isSwitchingLatexMode = true;

            try
            {
                // Gerenderte Formel verstecken
                renderedShape.Visible = Office.MsoTriState.msoFalse;

                // Ursprüngliches LaTeX-Feld wieder anzeigen
                sourceShape.Visible = Office.MsoTriState.msoTrue;

                // Als aktives LaTeX-Feld merken
                activeLatexField = sourceShape;

                // Shape auswählen
                sourceShape.Select(Office.MsoTriState.msoTrue);

                // Direkt Textbearbeitung aktivieren
                sourceShape.TextFrame.TextRange.Select();
            }
            finally
            {
                isSwitchingLatexMode = false;
            }
        }

        private void DeleteExistingRender(
            PowerPoint.Slide slide, int sourceShapeId)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                PowerPoint.Shape shape = slide.Shapes[i];

                try
                {
                    bool isLatexRender = shape.Tags["LatexRender"] == "true";

                    bool belongsToSource = shape.Tags["SourceShapeId"] == sourceShapeId.ToString();

                    if (isLatexRender && belongsToSource)
                    {
                        shape.Delete();
                    }
                }
                catch
                {
                    // Shape ist kein LaTeX-Render -> ignorieren
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