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
    public partial class SimpleControl : UserControl
    {
        public SimpleControl()
        {
            InitializeComponent();
        }

        

        private void btnGenTable_Click(object sender, EventArgs e)
        {
            //Word.Document document = Globals.ThisAddIn.Application.ActiveDocument; 
            //object start = 0;
            //object end = 0; 
            //Word.Range tableLocation = (Word.Range)document.Range(ref start, ref end);

            //tableLocation.Font.Size = 8;
            //object indentStyle = "Normal Indent";
            //tableLocation.set_Style(ref indentStyle);

            //document.Tables.Add(tableLocation, 
            //    Convert.ToInt32(txtNoOfRows.Text), Convert.ToInt32(txtNoOfRows.Text));



            object missing = System.Type.Missing;
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            Word.Table newTable = Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
            currentRange, Convert.ToInt32(txtNoOfRows.Text), Convert.ToInt32(txtNoOfRows.Text), 
            ref missing, ref missing);

            // Get all of the borders except for the diagonal borders.
            Word.Border[] borders = new Word.Border[6];
            borders[0] = newTable.Borders[Word.WdBorderType.wdBorderLeft];
            borders[1] = newTable.Borders[Word.WdBorderType.wdBorderRight];
            borders[2] = newTable.Borders[Word.WdBorderType.wdBorderTop];
            borders[3] = newTable.Borders[Word.WdBorderType.wdBorderBottom];
            borders[4] = newTable.Borders[Word.WdBorderType.wdBorderHorizontal];
            borders[5] = newTable.Borders[Word.WdBorderType.wdBorderVertical];

            // Format each of the borders.
            foreach (Word.Border border in borders)
            {
                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.Color = Word.WdColor.wdColorBlue;
            }
        }

        private void btnStrikethrough_Click(object sender, EventArgs e)
        {
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            //currentRange.Font.StrikeThrough = 1;
            currentRange.Font.Superscript = 1;
            //currentRange

        }

        private void btnSelectedText_Click(object sender, EventArgs e)
        {
            Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
            lblSelectedText.Text = currentRange.Text;
        }
    }
}
