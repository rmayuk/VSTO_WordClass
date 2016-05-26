using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VSTO_WordDemo1
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Globals.ThisDocument.Paragraphs[1].Range.Text = "Hello World!";
            //Globals.ThisAddIn.Application.Documents,1].Range.Text = "Hello World!";
        }
    }
}
