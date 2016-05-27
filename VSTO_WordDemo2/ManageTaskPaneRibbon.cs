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
            bool blnState =  ((RibbonToggleButton)sender).Checked;
            if (blnState == false)
                toggleButton1.Label = "Show Action Pane";
            else
                toggleButton1.Label = "Hide Action Pane";
            Globals.ThisAddIn.CustomTaskPanes[0].Visible = blnState;
        }

        private void tglBtnTextFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            bool blnState = ((RibbonToggleButton)sender).Checked;
            if (blnState == false)
                tglBtnTextFormatting.Label = "Show Text Pane";
            else
                tglBtnTextFormatting.Label = "Hide Text Pane";
            Globals.ThisAddIn.CustomTaskPanes[1].Visible = blnState;
        }
    }
}
