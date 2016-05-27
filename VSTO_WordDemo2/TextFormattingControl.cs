using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;


namespace VSTO_WordDemo2
{
    public partial class TextFormattingControl : UserControl
    {
        public TextFormattingControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(txtSimpleMsg.Text))
                { 
                    Word.Range range = Globals.ThisAddIn.Application.Selection.Range; 
                    range.InsertAfter(txtSimpleMsg.Text+"["+ range.Start.ToString()+"|"+ range.End.ToString()+"]");
                }
            }
            catch (COMException ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
            }
        }

        private void TextFormattingControl_Load(object sender, EventArgs e)
        {
            foreach (FontFamily font in System.Drawing.FontFamily.Families)
                cboFontNames.Items.Add(font.Name);
            for (int iRow = 8; iRow <=72; iRow++)
                cboFontSize.Items.Add(iRow.ToString());
        }

        private void btnSendFormattedText_Click(object sender, EventArgs e)
        {
            Word.Range range = Globals.ThisAddIn.Application.Selection.Range;
            range.Font.Size =Convert.ToInt32(cboFontSize.SelectedItem);
            range.Font.Name = cboFontNames.SelectedItem.ToString();
            range.InsertAfter(range.Text);
        }

        private void btnUndo_Click(object sender, EventArgs e)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            object numTimes3 = Convert.ToInt32( txtUndoNumber.Text);
            document.Undo(ref numTimes3);
        }

        private void btnFindSynonym_Click(object sender, EventArgs e)
        {
            lstSynonym.Items.Clear();
            var app = new Word.Application();
            var infosyn = app.SynonymInfo[txtSynonym.Text, Word.WdLanguageID.wdEnglishUS];
            foreach (var item in infosyn.MeaningList as Array)
                lstSynonym.Items.Add(item.ToString());
        }

        private void lstSynonym_DoubleClick(object sender, EventArgs e)
        {
            Word.Range range = Globals.ThisAddIn.Application.Selection.Range; 
            range.InsertAfter(lstSynonym.SelectedItem.ToString());
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            //        Word.Application.ActiveDocument.Range(
            //Word.Application.ActiveDocument.Content.Start,
            //Word.Application.ActiveDocument.Content.End).Select();

            //Globals.ThisAddIn.Application.ActiveDocument.Range(
            //Globals.ThisAddIn.Application.ActiveDocument.Content.Start,
            //Globals.ThisAddIn.Application.ActiveDocument.Content.End).Select();

            int start = Convert.ToInt32(txtSelectStart.Text);
            int end = Convert.ToInt32(txtSelectEnd.Text);
            Globals.ThisAddIn.Application.ActiveDocument.Range(start,end).Select();

        }
    }
}
