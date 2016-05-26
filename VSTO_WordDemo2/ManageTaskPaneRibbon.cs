using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace VSTO_WordDemo2
{
    public partial class ManageTaskPaneRibbon
    {
        private void ManageTaskPaneRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CustomTaskPanes[0].Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
