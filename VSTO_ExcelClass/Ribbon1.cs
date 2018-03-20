using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace VSTO_ExcelClass
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        object CustomDocumentProperties { get; }

        public Ribbon1()
        {
        }


        #region New Code

        void CreateProperties(string Classification)
        {
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Office.DocumentProperties)Globals.ThisAddIn.Application.ActiveWorkbook.BuiltinDocumentProperties;
            properties["Category"].Value = Classification;
        }

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

        #endregion




        #region Events Defination


        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("VSTO_ExcelClass.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {

            this.ribbon = ribbonUI;
        }

        public void ClassChange(Office.IRibbonControl control, string Test, int selectedIndex)
        {
            CreateProperties(Test);
            
            MessageBox.Show("This document has been classified " + ReadDocumentProperty("Category"),"Classified");

        }


        public string GetClassChange(Office.IRibbonControl control, int selectedIndex)
        {
            return ReadDocumentProperty("Category");

        }
        public string ddGetSelectedItemID(Office.IRibbonControl control)
        {
            if (ReadDocumentProperty("Category") != null)
            {
                return ReadDocumentProperty("Category");
            }
            else
                return "";
        }

        public string GetLabel(Office.IRibbonControl Control, int Index)
        {
            return "Item" + Index.ToString();
        }




        public string GetGuideState(Office.IRibbonControl control)
        {
            if (ReadDocumentProperty("Category") != null)
            {
                return ReadDocumentProperty("Category");
            }
            else
                return "";
        }

        public void AddHeader(Office.IRibbonControl control)
        {
            Excel.Workbook Workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (Workbook != null)
            {
                Microsoft.Office.Tools.Excel.Workbook vstoWorkbook = Globals.Factory.GetVstoObject(Workbook);

                Excel.Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
                if (worksheet != null)
                {
                    Microsoft.Office.Tools.Excel.Worksheet vstoSheet = Globals.Factory.GetVstoObject(worksheet);
                    vstoSheet.PageSetup.LeftHeader = ReadDocumentProperty("Category");
                }
            }
        }




        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }

    #endregion
}

