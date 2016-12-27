using ScalesData.Properties;

namespace ScalesData
{
    partial class ExportExcelForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.buttonDataBrowse = new System.Windows.Forms.Button();
            this.textBoxDbPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxDestination = new System.Windows.Forms.TextBox();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.startDatePicker = new System.Windows.Forms.DateTimePicker();
            this.btnExport = new System.Windows.Forms.Button();
            this.endDatePicker = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.checkBoxProtect = new System.Windows.Forms.CheckBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.tableLayoutPanel1);
            this.panel1.Controls.Add(this.startDatePicker);
            this.panel1.Controls.Add(this.btnExport);
            this.panel1.Controls.Add(this.endDatePicker);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(2, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(516, 270);
            this.panel1.TabIndex = 4;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.3658F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.6342F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 68F));
            this.tableLayoutPanel1.Controls.Add(this.buttonDataBrowse, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.textBoxDbPath, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBoxDestination, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.buttonBrowse, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(10, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 2;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(494, 68);
            this.tableLayoutPanel1.TabIndex = 12;
            // 
            // buttonDataBrowse
            // 
            this.buttonDataBrowse.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.buttonDataBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonDataBrowse.Location = new System.Drawing.Point(432, 39);
            this.buttonDataBrowse.Name = "buttonDataBrowse";
            this.buttonDataBrowse.Size = new System.Drawing.Size(54, 23);
            this.buttonDataBrowse.TabIndex = 14;
            this.buttonDataBrowse.Text = "...";
            this.buttonDataBrowse.UseVisualStyleBackColor = true;
            this.buttonDataBrowse.Click += new System.EventHandler(this.openSelectDestination);
            // 
            // textBoxDbPath
            // 
            this.textBoxDbPath.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.textBoxDbPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxDbPath.Location = new System.Drawing.Point(121, 41);
            this.textBoxDbPath.Name = "textBoxDbPath";
            this.textBoxDbPath.ReadOnly = true;
            this.textBoxDbPath.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.textBoxDbPath.Size = new System.Drawing.Size(295, 20);
            this.textBoxDbPath.TabIndex = 13;
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Destination Folder";
            // 
            // textBoxDestination
            // 
            this.textBoxDestination.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.textBoxDestination.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxDestination.Location = new System.Drawing.Point(121, 7);
            this.textBoxDestination.Name = "textBoxDestination";
            this.textBoxDestination.ReadOnly = true;
            this.textBoxDestination.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.textBoxDestination.Size = new System.Drawing.Size(295, 20);
            this.textBoxDestination.TabIndex = 5;
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.buttonBrowse.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonBrowse.Location = new System.Drawing.Point(433, 5);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(53, 23);
            this.buttonBrowse.TabIndex = 11;
            this.buttonBrowse.Text = "...";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.openSelectDestination);
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 44);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Select Data folder";
            // 
            // startDatePicker
            // 
            this.startDatePicker.CustomFormat = "dd/MM/yyyy HH:mm:ss";
            this.startDatePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.startDatePicker.Location = new System.Drawing.Point(10, 126);
            this.startDatePicker.Name = "startDatePicker";
            this.startDatePicker.ShowCheckBox = true;
            this.startDatePicker.Size = new System.Drawing.Size(200, 20);
            this.startDatePicker.TabIndex = 9;
            this.startDatePicker.Value = new System.DateTime(2016, 12, 15, 0, 0, 0, 0);
            // 
            // btnExport
            // 
            this.btnExport.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExport.Location = new System.Drawing.Point(215, 244);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(86, 23);
            this.btnExport.TabIndex = 8;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.onExportClick);
            // 
            // endDatePicker
            // 
            this.endDatePicker.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.endDatePicker.CustomFormat = "dd/MM/yyyy HH:mm:ss";
            this.endDatePicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.endDatePicker.Location = new System.Drawing.Point(304, 126);
            this.endDatePicker.Name = "endDatePicker";
            this.endDatePicker.ShowCheckBox = true;
            this.endDatePicker.Size = new System.Drawing.Size(200, 20);
            this.endDatePicker.TabIndex = 7;
            this.endDatePicker.Value = new System.DateTime(2016, 12, 15, 0, 0, 0, 0);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(419, 97);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(85, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Select end date:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 97);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Select start date:";
            // 
            // checkBoxProtect
            // 
            this.checkBoxProtect.AutoSize = true;
            this.checkBoxProtect.Location = new System.Drawing.Point(0, 21);
            this.checkBoxProtect.Name = "checkBoxProtect";
            this.checkBoxProtect.Size = new System.Drawing.Size(122, 17);
            this.checkBoxProtect.TabIndex = 13;
            this.checkBoxProtect.Text = "Protect by password";
            this.checkBoxProtect.UseVisualStyleBackColor = true;
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Enabled = false;
            this.textBoxPassword.Location = new System.Drawing.Point(155, 19);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.Size = new System.Drawing.Size(194, 20);
            this.textBoxPassword.TabIndex = 14;
            this.textBoxPassword.UseSystemPasswordChar = true;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.Controls.Add(this.checkBoxProtect);
            this.panel2.Controls.Add(this.textBoxPassword);
            this.panel2.Location = new System.Drawing.Point(87, 174);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(355, 49);
            this.panel2.TabIndex = 15;
            // 
            // ExportExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(518, 310);
            this.Controls.Add(this.panel1);
            this.Name = "ExportExcelForm";
            this.Text = "Export Excel";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnExport;
        protected System.Windows.Forms.DateTimePicker startDatePicker;
        protected System.Windows.Forms.DateTimePicker endDatePicker;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxDestination;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button buttonDataBrowse;
        private System.Windows.Forms.TextBox textBoxDbPath;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckBox checkBoxProtect;
        private System.Windows.Forms.TextBox textBoxPassword;
    }
}