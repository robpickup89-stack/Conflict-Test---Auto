
namespace Conflict_Test___Auto
{


    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.RichTextBox();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.ComboBox();
            this.groupBoxConnection = new System.Windows.Forms.GroupBox();
            this.groupBoxExport = new System.Windows.Forms.GroupBox();
            this.groupBoxTest = new System.Windows.Forms.GroupBox();
            this.statusLabel = new System.Windows.Forms.Label();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.groupBoxLog = new System.Windows.Forms.GroupBox();
            this.groupBoxMatrix = new System.Windows.Forms.GroupBox();
            this.conflictMatrix = new System.Windows.Forms.DataGridView();
            this.matrixLegend = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.conflictMatrix)).BeginInit();
            this.groupBoxConnection.SuspendLayout();
            this.groupBoxExport.SuspendLayout();
            this.groupBoxTest.SuspendLayout();
            this.groupBoxLog.SuspendLayout();
            this.groupBoxMatrix.SuspendLayout();
            this.SuspendLayout();
            //
            // groupBoxConnection
            //
            this.groupBoxConnection.Controls.Add(this.label1);
            this.groupBoxConnection.Controls.Add(this.textBox1);
            this.groupBoxConnection.Controls.Add(this.label3);
            this.groupBoxConnection.Controls.Add(this.numericUpDown2);
            this.groupBoxConnection.Controls.Add(this.button6);
            this.groupBoxConnection.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBoxConnection.Location = new System.Drawing.Point(12, 12);
            this.groupBoxConnection.Name = "groupBoxConnection";
            this.groupBoxConnection.Size = new System.Drawing.Size(280, 85);
            this.groupBoxConnection.TabIndex = 0;
            this.groupBoxConnection.TabStop = false;
            this.groupBoxConnection.Text = "Connection Settings";
            //
            // groupBoxExport
            //
            this.groupBoxExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxExport.Controls.Add(this.button4);
            this.groupBoxExport.Controls.Add(this.button5);
            this.groupBoxExport.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBoxExport.Location = new System.Drawing.Point(298, 12);
            this.groupBoxExport.Name = "groupBoxExport";
            this.groupBoxExport.Size = new System.Drawing.Size(140, 85);
            this.groupBoxExport.TabIndex = 1;
            this.groupBoxExport.TabStop = false;
            this.groupBoxExport.Text = "Export Results";
            //
            // groupBoxTest
            //
            this.groupBoxTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxTest.Controls.Add(this.button1);
            this.groupBoxTest.Controls.Add(this.button3);
            this.groupBoxTest.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBoxTest.Location = new System.Drawing.Point(444, 12);
            this.groupBoxTest.Name = "groupBoxTest";
            this.groupBoxTest.Size = new System.Drawing.Size(128, 85);
            this.groupBoxTest.TabIndex = 2;
            this.groupBoxTest.TabStop = false;
            this.groupBoxTest.Text = "Test Control";
            //
            // button1
            //
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(144)))), ((int)(((byte)(238)))), ((int)(((byte)(144)))));
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(60)))), ((int)(((byte)(179)))), ((int)(((byte)(113)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(100)))), ((int)(((byte)(0)))));
            this.button1.Location = new System.Drawing.Point(12, 22);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(104, 26);
            this.button1.TabIndex = 4;
            this.button1.Text = "\u25B6 Start Test";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            //
            // label1
            //
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(10, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "IP Address:";
            //
            // numericUpDown2
            //
            this.numericUpDown2.Location = new System.Drawing.Point(78, 52);
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(55, 23);
            this.numericUpDown2.TabIndex = 2;
            this.numericUpDown2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numericUpDown2.ValueChanged += new System.EventHandler(this.numericUpDown2_ValueChanged);
            //
            // label3
            //
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(10, 54);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 15);
            this.label3.TabIndex = 7;
            this.label3.Text = "COM Port:";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            //
            // button3
            //
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(182)))), ((int)(((byte)(182)))));
            this.button3.Enabled = false;
            this.button3.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(53)))), ((int)(((byte)(69)))));
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(139)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.button3.Location = new System.Drawing.Point(12, 52);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(104, 26);
            this.button3.TabIndex = 5;
            this.button3.Text = "\u25A0 Stop Test";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            //
            // splitContainer
            //
            this.splitContainer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer.Location = new System.Drawing.Point(12, 103);
            this.splitContainer.Name = "splitContainer";
            //
            // splitContainer.Panel1
            //
            this.splitContainer.Panel1.Controls.Add(this.groupBoxLog);
            this.splitContainer.Panel1MinSize = 250;
            //
            // splitContainer.Panel2
            //
            this.splitContainer.Panel2.Controls.Add(this.groupBoxMatrix);
            this.splitContainer.Panel2MinSize = 300;
            this.splitContainer.Size = new System.Drawing.Size(760, 291);
            this.splitContainer.SplitterDistance = 400;
            this.splitContainer.TabIndex = 13;
            //
            // groupBoxLog
            //
            this.groupBoxLog.Controls.Add(this.textBox2);
            this.groupBoxLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBoxLog.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBoxLog.Location = new System.Drawing.Point(0, 0);
            this.groupBoxLog.Name = "groupBoxLog";
            this.groupBoxLog.Padding = new System.Windows.Forms.Padding(6);
            this.groupBoxLog.Size = new System.Drawing.Size(400, 291);
            this.groupBoxLog.TabIndex = 0;
            this.groupBoxLog.TabStop = false;
            this.groupBoxLog.Text = "Test Log";
            //
            // textBox2
            //
            this.textBox2.BackColor = System.Drawing.Color.White;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.textBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox2.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.textBox2.Location = new System.Drawing.Point(6, 22);
            this.textBox2.Margin = new System.Windows.Forms.Padding(10);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(388, 263);
            this.textBox2.TabIndex = 10;
            this.textBox2.Text = "";
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged_1);
            //
            // groupBoxMatrix
            //
            this.groupBoxMatrix.Controls.Add(this.conflictMatrix);
            this.groupBoxMatrix.Controls.Add(this.matrixLegend);
            this.groupBoxMatrix.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBoxMatrix.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.groupBoxMatrix.Location = new System.Drawing.Point(0, 0);
            this.groupBoxMatrix.Name = "groupBoxMatrix";
            this.groupBoxMatrix.Padding = new System.Windows.Forms.Padding(6);
            this.groupBoxMatrix.Size = new System.Drawing.Size(356, 291);
            this.groupBoxMatrix.TabIndex = 0;
            this.groupBoxMatrix.TabStop = false;
            this.groupBoxMatrix.Text = "Conflict Matrix (From \u2192 To)";
            //
            // conflictMatrix
            //
            this.conflictMatrix.AllowUserToAddRows = false;
            this.conflictMatrix.AllowUserToDeleteRows = false;
            this.conflictMatrix.AllowUserToResizeColumns = false;
            this.conflictMatrix.AllowUserToResizeRows = false;
            this.conflictMatrix.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.conflictMatrix.BackgroundColor = System.Drawing.Color.White;
            this.conflictMatrix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.conflictMatrix.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.conflictMatrix.EnableHeadersVisualStyles = false;
            this.conflictMatrix.Location = new System.Drawing.Point(6, 22);
            this.conflictMatrix.Name = "conflictMatrix";
            this.conflictMatrix.ReadOnly = true;
            this.conflictMatrix.RowHeadersWidth = 35;
            this.conflictMatrix.RowTemplate.Height = 25;
            this.conflictMatrix.Size = new System.Drawing.Size(344, 239);
            this.conflictMatrix.TabIndex = 0;
            //
            // matrixLegend
            //
            this.matrixLegend.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.matrixLegend.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.matrixLegend.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.matrixLegend.Location = new System.Drawing.Point(6, 264);
            this.matrixLegend.Name = "matrixLegend";
            this.matrixLegend.Size = new System.Drawing.Size(344, 20);
            this.matrixLegend.TabIndex = 1;
            this.matrixLegend.Text = "\u2713 Pass  |  \u2717 Fail  |  \u25CB Pending  |  \u25CF Running  |  - No Conflict";
            this.matrixLegend.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // button4
            //
            this.button4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.button4.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(180)))), ((int)(((byte)(180)))));
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button4.Location = new System.Drawing.Point(12, 22);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(116, 26);
            this.button4.TabIndex = 8;
            this.button4.Text = "\uD83D\uDCC4 Export RTF";
            this.button4.UseVisualStyleBackColor = false;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            //
            // button5
            //
            this.button5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.button5.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(180)))), ((int)(((byte)(180)))), ((int)(((byte)(180)))));
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button5.Location = new System.Drawing.Point(12, 52);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(116, 26);
            this.button5.TabIndex = 9;
            this.button5.Text = "\uD83D\uDCC4 Export PDF";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            //
            // button6
            //
            this.button6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(230)))), ((int)(((byte)(240)))), ((int)(((byte)(255)))));
            this.button6.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(100)))), ((int)(((byte)(149)))), ((int)(((byte)(237)))));
            this.button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button6.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.button6.Location = new System.Drawing.Point(139, 50);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(130, 26);
            this.button6.TabIndex = 7;
            this.button6.Text = "\uD83D\uDD0D Auto Detect";
            this.button6.UseVisualStyleBackColor = false;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            //
            // textBox1
            //
            this.textBox1.FormattingEnabled = true;
            this.textBox1.Items.AddRange(new object[] {
            "10.164.95.198",
            "10.164.95.199",
            "10.164.95.200",
            "10.164.95.201",
            "10.164.95.202"});
            this.textBox1.Location = new System.Drawing.Point(78, 21);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(130, 23);
            this.textBox1.TabIndex = 11;
            //
            // statusLabel
            //
            this.statusLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.statusLabel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(249)))), ((int)(((byte)(250)))));
            this.statusLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.statusLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.statusLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(80)))), ((int)(((byte)(87)))));
            this.statusLabel.Location = new System.Drawing.Point(12, 400);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Padding = new System.Windows.Forms.Padding(5, 0, 0, 0);
            this.statusLabel.Size = new System.Drawing.Size(776, 24);
            this.statusLabel.TabIndex = 12;
            this.statusLabel.Text = "Ready";
            this.statusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // Form1
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(250)))), ((int)(((byte)(250)))), ((int)(((byte)(250)))));
            this.ClientSize = new System.Drawing.Size(800, 431);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.groupBoxTest);
            this.Controls.Add(this.groupBoxExport);
            this.Controls.Add(this.groupBoxConnection);
            this.Controls.Add(this.splitContainer);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(816, 470);
            this.Name = "Form1";
            this.Text = "Auto Conflict Test - Swarco";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.conflictMatrix)).EndInit();
            this.groupBoxConnection.ResumeLayout(false);
            this.groupBoxConnection.PerformLayout();
            this.groupBoxExport.ResumeLayout(false);
            this.groupBoxTest.ResumeLayout(false);
            this.groupBoxLog.ResumeLayout(false);
            this.groupBoxMatrix.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.RichTextBox textBox2;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.ComboBox textBox1;
        private System.Windows.Forms.GroupBox groupBoxConnection;
        private System.Windows.Forms.GroupBox groupBoxExport;
        private System.Windows.Forms.GroupBox groupBoxTest;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.SplitContainer splitContainer;
        private System.Windows.Forms.GroupBox groupBoxLog;
        private System.Windows.Forms.GroupBox groupBoxMatrix;
        private System.Windows.Forms.DataGridView conflictMatrix;
        private System.Windows.Forms.Label matrixLegend;
    }
}

