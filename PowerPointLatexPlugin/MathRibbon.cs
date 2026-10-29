using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLatexPlugin
{
    public partial class MathRibbon
    {
        private void MathRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void tglInsertLatex_Click(object sender, RibbonControlEventArgs e)
        {
            if (tglInsertLatex.Checked)
            {
                Globals.ThisAddIn.StartLatexFieldDrawing();
            }
        }

        public void ResetInsertLatexToggle()
        {
            tglInsertLatex.Checked = false;
        }
    }
}
