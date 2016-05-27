namespace VSTO_WordDemo2
{
    partial class TextFormattingControl
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
            this.txtSimpleMsg = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnSendFormattedText = new System.Windows.Forms.Button();
            this.cboFontSize = new System.Windows.Forms.ComboBox();
            this.cboFontNames = new System.Windows.Forms.ComboBox();
            this.btnUndo = new System.Windows.Forms.Button();
            this.txtUndoNumber = new System.Windows.Forms.TextBox();
            this.lstSynonym = new System.Windows.Forms.ListBox();
            this.btnFindSynonym = new System.Windows.Forms.Button();
            this.txtSynonym = new System.Windows.Forms.TextBox();
            this.btnSelect = new System.Windows.Forms.Button();
            this.txtSelectStart = new System.Windows.Forms.TextBox();
            this.txtSelectEnd = new System.Windows.Forms.TextBox();
            this.txtSelectedDisplay = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtSimpleMsg
            // 
            this.txtSimpleMsg.Location = new System.Drawing.Point(12, 3);
            this.txtSimpleMsg.Name = "txtSimpleMsg";
            this.txtSimpleMsg.Size = new System.Drawing.Size(111, 20);
            this.txtSimpleMsg.TabIndex = 3;
            this.txtSimpleMsg.Text = "Welcome";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(129, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(70, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Send Msg";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSendFormattedText
            // 
            this.btnSendFormattedText.Location = new System.Drawing.Point(81, 62);
            this.btnSendFormattedText.Name = "btnSendFormattedText";
            this.btnSendFormattedText.Size = new System.Drawing.Size(124, 23);
            this.btnSendFormattedText.TabIndex = 2;
            this.btnSendFormattedText.Text = "Send Formatted Text";
            this.btnSendFormattedText.UseVisualStyleBackColor = true;
            this.btnSendFormattedText.Click += new System.EventHandler(this.btnSendFormattedText_Click);
            // 
            // cboFontSize
            // 
            this.cboFontSize.FormattingEnabled = true;
            this.cboFontSize.Location = new System.Drawing.Point(7, 64);
            this.cboFontSize.Name = "cboFontSize";
            this.cboFontSize.Size = new System.Drawing.Size(59, 21);
            this.cboFontSize.TabIndex = 4;
            // 
            // cboFontNames
            // 
            this.cboFontNames.FormattingEnabled = true;
            this.cboFontNames.Location = new System.Drawing.Point(7, 37);
            this.cboFontNames.Name = "cboFontNames";
            this.cboFontNames.Size = new System.Drawing.Size(202, 21);
            this.cboFontNames.TabIndex = 4;
            // 
            // btnUndo
            // 
            this.btnUndo.Location = new System.Drawing.Point(81, 99);
            this.btnUndo.Name = "btnUndo";
            this.btnUndo.Size = new System.Drawing.Size(124, 23);
            this.btnUndo.TabIndex = 2;
            this.btnUndo.Text = "Undo";
            this.btnUndo.UseVisualStyleBackColor = true;
            this.btnUndo.Click += new System.EventHandler(this.btnUndo_Click);
            // 
            // txtUndoNumber
            // 
            this.txtUndoNumber.Location = new System.Drawing.Point(7, 101);
            this.txtUndoNumber.Name = "txtUndoNumber";
            this.txtUndoNumber.Size = new System.Drawing.Size(59, 20);
            this.txtUndoNumber.TabIndex = 3;
            this.txtUndoNumber.Text = "2";
            // 
            // lstSynonym
            // 
            this.lstSynonym.FormattingEnabled = true;
            this.lstSynonym.Location = new System.Drawing.Point(7, 163);
            this.lstSynonym.Name = "lstSynonym";
            this.lstSynonym.Size = new System.Drawing.Size(192, 186);
            this.lstSynonym.TabIndex = 5;
            this.lstSynonym.DoubleClick += new System.EventHandler(this.lstSynonym_DoubleClick);
            // 
            // btnFindSynonym
            // 
            this.btnFindSynonym.Location = new System.Drawing.Point(113, 134);
            this.btnFindSynonym.Name = "btnFindSynonym";
            this.btnFindSynonym.Size = new System.Drawing.Size(86, 23);
            this.btnFindSynonym.TabIndex = 6;
            this.btnFindSynonym.Text = "Find Synonym";
            this.btnFindSynonym.UseVisualStyleBackColor = true;
            this.btnFindSynonym.Click += new System.EventHandler(this.btnFindSynonym_Click);
            // 
            // txtSynonym
            // 
            this.txtSynonym.Location = new System.Drawing.Point(7, 134);
            this.txtSynonym.Name = "txtSynonym";
            this.txtSynonym.Size = new System.Drawing.Size(100, 20);
            this.txtSynonym.TabIndex = 7;
            this.txtSynonym.Text = "almost";
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(147, 374);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(52, 23);
            this.btnSelect.TabIndex = 2;
            this.btnSelect.Text = "Select";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // txtSelectStart
            // 
            this.txtSelectStart.Location = new System.Drawing.Point(7, 374);
            this.txtSelectStart.Name = "txtSelectStart";
            this.txtSelectStart.Size = new System.Drawing.Size(59, 20);
            this.txtSelectStart.TabIndex = 3;
            this.txtSelectStart.Text = "2";
            // 
            // txtSelectEnd
            // 
            this.txtSelectEnd.Location = new System.Drawing.Point(72, 374);
            this.txtSelectEnd.Name = "txtSelectEnd";
            this.txtSelectEnd.Size = new System.Drawing.Size(59, 20);
            this.txtSelectEnd.TabIndex = 3;
            this.txtSelectEnd.Text = "5";
            // 
            // txtSelectedDisplay
            // 
            this.txtSelectedDisplay.Location = new System.Drawing.Point(7, 413);
            this.txtSelectedDisplay.Name = "txtSelectedDisplay";
            this.txtSelectedDisplay.Size = new System.Drawing.Size(192, 20);
            this.txtSelectedDisplay.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 397);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Selected Text Range";
            // 
            // TextFormattingControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtSynonym);
            this.Controls.Add(this.btnFindSynonym);
            this.Controls.Add(this.lstSynonym);
            this.Controls.Add(this.cboFontNames);
            this.Controls.Add(this.cboFontSize);
            this.Controls.Add(this.txtSelectEnd);
            this.Controls.Add(this.txtSelectStart);
            this.Controls.Add(this.txtUndoNumber);
            this.Controls.Add(this.txtSelectedDisplay);
            this.Controls.Add(this.txtSimpleMsg);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.btnUndo);
            this.Controls.Add(this.btnSendFormattedText);
            this.Controls.Add(this.button1);
            this.Name = "TextFormattingControl";
            this.Size = new System.Drawing.Size(254, 451);
            this.Load += new System.EventHandler(this.TextFormattingControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSimpleMsg;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnSendFormattedText;
        private System.Windows.Forms.ComboBox cboFontSize;
        private System.Windows.Forms.ComboBox cboFontNames;
        private System.Windows.Forms.Button btnUndo;
        private System.Windows.Forms.TextBox txtUndoNumber;
        private System.Windows.Forms.ListBox lstSynonym;
        private System.Windows.Forms.Button btnFindSynonym;
        private System.Windows.Forms.TextBox txtSynonym;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.TextBox txtSelectStart;
        private System.Windows.Forms.TextBox txtSelectEnd;
        private System.Windows.Forms.TextBox txtSelectedDisplay;
        private System.Windows.Forms.Label label1;
    }
}
