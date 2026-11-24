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

        private System.Windows.Forms.Timer latexPreviewTimer;
        private PowerPoint.Shape activeLatexPreview = null;

        private float lastPreviewLeft;
        private float lastPreviewTop;
        private float lastPreviewWidth;
        private float lastPreviewHeight;

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

            latexPreviewTimer = new System.Windows.Forms.Timer();

            latexPreviewTimer.Interval = 50;

            latexPreviewTimer.Tick += LatexPreviewTimer_Tick;
        }


        public void StartLatexFieldDrawing()
        {
            latexFieldDrawer.Start();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            latexPreviewTimer?.Stop();
            latexPreviewTimer?.Dispose();

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
                    latexPreviewTimer.Stop();
                    activeLatexPreview = null;

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

                        DeleteExistingPreview(slide, activeLatexField.Id);

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

                        FitShapeIntoLatexField(renderedFormula, activeLatexField);

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

                // halbtransparente Formel-Vorschau erstellen
                activeLatexPreview =
                    CreateLatexPreview(sourceShape);

                // Aktuelle Position/Größe merken
                lastPreviewLeft = sourceShape.Left;
                lastPreviewTop = sourceShape.Top;
                lastPreviewWidth = sourceShape.Width;
                lastPreviewHeight = sourceShape.Height;

                // Live-Synchronisierung starten
                if (activeLatexPreview != null)
                {
                    latexPreviewTimer.Start();
                }

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

        private void LatexPreviewTimer_Tick(object sender, EventArgs e)
        {
            if (activeLatexField == null ||
                activeLatexPreview == null)
            {
                return;
            }

            try
            {
                float currentLeft =
                    activeLatexField.Left;

                float currentTop =
                    activeLatexField.Top;

                float currentWidth =
                    activeLatexField.Width;

                float currentHeight =
                    activeLatexField.Height;

                bool changed =
                    Math.Abs(currentLeft - lastPreviewLeft) > 0.1f ||
                    Math.Abs(currentTop - lastPreviewTop) > 0.1f ||
                    Math.Abs(currentWidth - lastPreviewWidth) > 0.1f ||
                    Math.Abs(currentHeight - lastPreviewHeight) > 0.1f;

                if (!changed)
                    return;

                FitShapeIntoLatexField(
                    activeLatexPreview,
                    activeLatexField
                );

                lastPreviewLeft = currentLeft;
                lastPreviewTop = currentTop;
                lastPreviewWidth = currentWidth;
                lastPreviewHeight = currentHeight;
            }
            catch
            {
                latexPreviewTimer.Stop();
                activeLatexPreview = null;
            }
        }

        private PowerPoint.Shape CreateLatexPreview(
            PowerPoint.Shape sourceShape)
        {
            string latexSource =
                sourceShape.Tags["LatexSource"];

            if (string.IsNullOrEmpty(latexSource))
                return null;

            PowerPoint.Slide slide =
                this.Application.ActiveWindow.View.Slide;

            // Erst alte Preview dieses Feldes entfernen
            DeleteExistingPreview(
                slide,
                sourceShape.Id
            );

            string previewPath =
                LatexRenderer.RenderToSvg(
                    latexSource,
                    0.35
                );

            PowerPoint.Shape preview =
                slide.Shapes.AddPicture(
                    previewPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    sourceShape.Left,
                    sourceShape.Top,
                    -1,
                    -1
                );

            FitShapeIntoLatexField(
                preview,
                sourceShape
            );

            preview.Tags.Add(
                "LatexPreview",
                "true"
            );

            preview.Tags.Add(
                "SourceShapeId",
                sourceShape.Id.ToString()
            );

            preview.ZOrder(
                Office.MsoZOrderCmd.msoSendBackward
            );

            return preview;
        }

        private void DeleteExistingPreview(PowerPoint.Slide slide, int sourceShapeId)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                PowerPoint.Shape shape = slide.Shapes[i];

                try
                {
                    bool isPreview = shape.Tags["LatexPreview"] == "true";

                    bool belongsToSource = shape.Tags["SourceShapeId"] == sourceShapeId.ToString();

                    if(isPreview && belongsToSource)
                    {
                        shape.Delete();
                    }
                }
                catch
                {
                }     
            }
        }

        private void FitShapeIntoLatexField(
            PowerPoint.Shape shape,
            PowerPoint.Shape sourceShape)
        {
            float maxWidth = sourceShape.Width;
            float maxHeight = sourceShape.Height;

            float originalWidth = shape.Width;
            float originalHeight = shape.Height;

            float widthScale =
                maxWidth / originalWidth;

            float heightScale =
                maxHeight / originalHeight;

            float scaleFactor =
                Math.Min(widthScale, heightScale);

            // Wichtig:
            // Breite UND Höhe explizit proportional setzen
            shape.LockAspectRatio =
                Office.MsoTriState.msoFalse;

            shape.Width =
                originalWidth * scaleFactor;

            shape.Height =
                originalHeight * scaleFactor;

            // Im LaTeX-Feld zentrieren
            shape.Left =
                sourceShape.Left +
                (maxWidth - shape.Width) / 2;

            shape.Top =
                sourceShape.Top +
                (maxHeight - shape.Height) / 2;
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