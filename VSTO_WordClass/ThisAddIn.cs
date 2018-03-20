using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using static VSTO_WordClass.Ribbon1;
using System.Threading;

namespace VSTO_WordClass
{

    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender,
                                        System.EventArgs e)
        {
            // attach the save handler
            Globals.ThisAddIn.Application.DocumentBeforeSave += Application_DocumentBeforeSave;

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
        /// 

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        public string ReadDocumentProperty(string propertyName)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties;

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

        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties;
            if (ReadDocumentProperty("Category") == null)
            {
                MessageBox.Show("You are about to save a document without a classification. " + "\r\n" +
                                "If this document is the property of virtualDCS it must me classified" + "\r\n" +
                                "Please cancel your save and click on the classification tab on the ribbon.",
                                "                                    ------- WARNING -------", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        }

    }
