using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace VSTO_WordDemo2
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane; 


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _taskPane = this.CustomTaskPanes.Add(new SimpleControl(), "Syed's Docu.Builder");
            _taskPane.Visible = true;
            _taskPane.VisibleChanged +=new EventHandler(_taskPane_VisibleChanged);
        }

        private void _taskPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked = _taskPane.Visible;
        } 
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
