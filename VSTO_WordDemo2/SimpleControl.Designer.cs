namespace VSTO_WordDemo2
{
    partial class SimpleControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.txtNoOfRows = new System.Windows.Forms.TextBox();
            this.txtNoOfCols = new System.Windows.Forms.TextBox();
            this.btnGenTable = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnStrikethrough = new System.Windows.Forms.Button();
            this.lblSelectedText = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnSelectedText = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(39, 34);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(70, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Send Msg";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(8, 8);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(128, 20);
            this.textBox1.TabIndex = 1;
            // 
            // txtNoOfRows
            // 
            this.txtNoOfRows.Location = new System.Drawing.Point(9, 98);
            this.txtNoOfRows.Name = "txtNoOfRows";
            this.txtNoOfRows.Size = new System.Drawing.Size(100, 20);
            this.txtNoOfRows.TabIndex = 2;
            this.txtNoOfRows.Text = "3";
            // 
            // txtNoOfCols
            // 
            this.txtNoOfCols.Location = new System.Drawing.Point(9, 146);
            this.txtNoOfCols.Name = "txtNoOfCols";
            this.txtNoOfCols.Size = new System.Drawing.Size(100, 20);
            this.txtNoOfCols.TabIndex = 3;
            this.txtNoOfCols.Text = "4";
            // 
            // btnGenTable
            // 
            this.btnGenTable.Location = new System.Drawing.Point(9, 184);
            this.btnGenTable.Name = "btnGenTable";
            this.btnGenTable.Size = new System.Drawing.Size(75, 23);
            this.btnGenTable.TabIndex = 4;
            this.btnGenTable.Text = "Gen. Table";
            this.btnGenTable.UseVisualStyleBackColor = true;
            this.btnGenTable.Click += new System.EventHandler(this.btnGenTable_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "No of Rows";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 126);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "No of Columns";
            // 
            // btnStrikethrough
            // 
            this.btnStrikethrough.Location = new System.Drawing.Point(22, 350);
            this.btnStrikethrough.Name = "btnStrikethrough";
            this.btnStrikethrough.Size = new System.Drawing.Size(87, 23);
            this.btnStrikethrough.TabIndex = 6;
            this.btnStrikethrough.Text = "Strikethrough";
            this.btnStrikethrough.UseVisualStyleBackColor = true;
            this.btnStrikethrough.Click += new System.EventHandler(this.btnStrikethrough_Click);
            // 
            // lblSelectedText
            // 
            this.lblSelectedText.AutoSize = true;
            this.lblSelectedText.Location = new System.Drawing.Point(6, 251);
            this.lblSelectedText.Name = "lblSelectedText";
            this.lblSelectedText.Size = new System.Drawing.Size(91, 13);
            this.lblSelectedText.TabIndex = 5;
            this.lblSelectedText.Text = "______________";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(20, 228);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Selected Text";
            // 
            // btnSelectedText
            // 
            this.btnSelectedText.Location = new System.Drawing.Point(10, 276);
            this.btnSelectedText.Name = "btnSelectedText";
            this.btnSelectedText.Size = new System.Drawing.Size(109, 23);
            this.btnSelectedText.TabIndex = 6;
            this.btnSelectedText.Text = "Get Selected Text";
            this.btnSelectedText.UseVisualStyleBackColor = true;
            this.btnSelectedText.Click += new System.EventHandler(this.btnSelectedText_Click);
            // 
            // SimpleControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnSelectedText);
            this.Controls.Add(this.btnStrikethrough);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblSelectedText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnGenTable);
            this.Controls.Add(this.txtNoOfCols);
            this.Controls.Add(this.txtNoOfRows);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Name = "SimpleControl";
            this.Size = new System.Drawing.Size(150, 472);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox txtNoOfRows;
        private System.Windows.Forms.TextBox txtNoOfCols;
        private System.Windows.Forms.Button btnGenTable;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnStrikethrough;
        private System.Windows.Forms.Label lblSelectedText;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnSelectedText;
    }
}
