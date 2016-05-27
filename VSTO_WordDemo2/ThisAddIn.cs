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
        private Microsoft.Office.Tools.CustomTaskPane _textPane;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region Action Pane
            _taskPane = this.CustomTaskPanes.Add(new SimpleControl(), "AIT Action Pane");
            _taskPane.Width = 300;
            _taskPane.Visible = false;
            _taskPane.VisibleChanged +=new EventHandler(_taskPane_VisibleChanged);
            #endregion

            #region Text Formatting Pane
            _textPane = this.CustomTaskPanes.Add(new TextFormattingControl(), "AIT Text Pane");
            _textPane.Width = 300;
            _textPane.Visible = true;
            _textPane.VisibleChanged += new EventHandler(_textPane_VisibleChanged);
            #endregion
        }

        private void _taskPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked = _taskPane.Visible;
        }
        private void _textPane_VisibleChanged(object sender, EventArgs e)
        {
            Globals.Ribbons.ManageTaskPaneRibbon.tglBtnTextFormatting.Checked = _textPane.Visible;
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
