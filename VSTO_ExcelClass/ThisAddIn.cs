using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace VSTO_ExcelClass
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender,
                                        System.EventArgs e)
        {
            // attach the save handler

            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(oWord_DocumentBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
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

        #region new code

        public string ReadDocumentProperty(string propertyName)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)Globals.ThisAddIn.Application.ActiveWorkbook.BuiltinDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    if (prop.Value != null)
                        return prop.Value.ToString();
                }
            }
            return null;
        }


        void oWord_DocumentBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)Globals.ThisAddIn.Application.ActiveWorkbook.BuiltinDocumentProperties; 
            if (ReadDocumentProperty("Category") == null)
            {
                MessageBox.Show("You are about to save a document without a classification. " + "\r\n" +
                                "If this document is the property of virtualDCS it must me classified" + "\r\n" +
                                "Please cancel your save and click on the classification tab on the ribbon.",
                                "                                    ------- WARNING -------", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

    }


        #endregion


    }
