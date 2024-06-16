namespace ExcelSpliter3
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.Worksheets = new System.Windows.Forms.ComboBox();
            this.Headers = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Sheets2Files = new System.Windows.Forms.Button();
            this.Sheet2File = new System.Windows.Forms.Button();
            this.sendMail = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.MailPart = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.mailBody = new System.Windows.Forms.RichTextBox();
            this.mailSubject = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.logBox = new System.Windows.Forms.ListBox();
            this.ExportTo = new System.Windows.Forms.CheckedListBox();
            this.Autofit = new System.Windows.Forms.CheckBox();
            this.RemoveDuplicates = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.workBookPassword = new System.Windows.Forms.TextBox();
            this.setWorkbookPassword = new System.Windows.Forms.CheckBox();
            this.freezTopRow = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(777, 9);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(146, 23);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "باز کردن فایل اکسل";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // Worksheets
            // 
            this.Worksheets.Enabled = false;
            this.Worksheets.FormattingEnabled = true;
            this.Worksheets.Location = new System.Drawing.Point(777, 38);
            this.Worksheets.Name = "Worksheets";
            this.Worksheets.Size = new System.Drawing.Size(146, 21);
            this.Worksheets.TabIndex = 1;
            this.Worksheets.SelectedIndexChanged += new System.EventHandler(this.Worksheets_SelectedIndexChanged);
            // 
            // Headers
            // 
            this.Headers.Enabled = false;
            this.Headers.FormattingEnabled = true;
            this.Headers.Location = new System.Drawing.Point(777, 65);
            this.Headers.Name = "Headers";
            this.Headers.Size = new System.Drawing.Size(146, 21);
            this.Headers.TabIndex = 2;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dataGridView1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox1.Location = new System.Drawing.Point(0, 358);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox1.Size = new System.Drawing.Size(930, 331);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "پیش نمایش";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 16);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(924, 312);
            this.dataGridView1.TabIndex = 4;
            // 
            // Sheets2Files
            // 
            this.Sheets2Files.Enabled = false;
            this.Sheets2Files.Location = new System.Drawing.Point(618, 9);
            this.Sheets2Files.Name = "Sheets2Files";
            this.Sheets2Files.Size = new System.Drawing.Size(153, 23);
            this.Sheets2Files.TabIndex = 5;
            this.Sheets2Files.Text = "شکستن همه سربرگ ها";
            this.Sheets2Files.UseVisualStyleBackColor = true;
            this.Sheets2Files.Click += new System.EventHandler(this.Sheets2Files_Click);
            // 
            // Sheet2File
            // 
            this.Sheet2File.Enabled = false;
            this.Sheet2File.Location = new System.Drawing.Point(618, 38);
            this.Sheet2File.Name = "Sheet2File";
            this.Sheet2File.Size = new System.Drawing.Size(153, 23);
            this.Sheet2File.TabIndex = 6;
            this.Sheet2File.Text = "شکستن یک سربرگ";
            this.Sheet2File.UseVisualStyleBackColor = true;
            this.Sheet2File.Click += new System.EventHandler(this.Sheet2File_Click);
            // 
            // sendMail
            // 
            this.sendMail.AutoSize = true;
            this.sendMail.Enabled = false;
            this.sendMail.Location = new System.Drawing.Point(340, 19);
            this.sendMail.Name = "sendMail";
            this.sendMail.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.sendMail.Size = new System.Drawing.Size(116, 17);
            this.sendMail.TabIndex = 8;
            this.sendMail.Text = "ارسال ایمیل همزمان";
            this.sendMail.UseVisualStyleBackColor = true;
            this.sendMail.CheckedChanged += new System.EventHandler(this.sendMail_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.MailPart);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.sendMail);
            this.groupBox2.Controls.Add(this.mailBody);
            this.groupBox2.Controls.Add(this.mailSubject);
            this.groupBox2.Controls.Add(this.pictureBox1);
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox2.Size = new System.Drawing.Size(465, 346);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "تنظیمات ایمیل";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(393, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "موضوع ایمیل";
            // 
            // MailPart
            // 
            this.MailPart.Location = new System.Drawing.Point(100, 19);
            this.MailPart.Name = "MailPart";
            this.MailPart.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.MailPart.Size = new System.Drawing.Size(100, 20);
            this.MailPart.TabIndex = 13;
            this.MailPart.Text = "@agri-bank.com";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(408, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(51, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "بدنه ایمیل";
            // 
            // mailBody
            // 
            this.mailBody.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mailBody.Location = new System.Drawing.Point(6, 96);
            this.mailBody.Name = "mailBody";
            this.mailBody.Size = new System.Drawing.Size(452, 244);
            this.mailBody.TabIndex = 10;
            this.mailBody.Text = "";
            // 
            // mailSubject
            // 
            this.mailSubject.Location = new System.Drawing.Point(97, 55);
            this.mailSubject.Name = "mailSubject";
            this.mailSubject.Size = new System.Drawing.Size(359, 20);
            this.mailSubject.TabIndex = 9;
            this.mailSubject.Text = "موضوع ایمیل را وارد کنید";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(6, 10);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(87, 67);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 8;
            this.pictureBox1.TabStop = false;
            // 
            // logBox
            // 
            this.logBox.FormattingEnabled = true;
            this.logBox.Items.AddRange(new object[] {
            "برنامه نویس: شهاب صادقی",
            "لایسنس: رایگان"});
            this.logBox.Location = new System.Drawing.Point(618, 185);
            this.logBox.Name = "logBox";
            this.logBox.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.logBox.Size = new System.Drawing.Size(300, 173);
            this.logBox.TabIndex = 10;
            // 
            // ExportTo
            // 
            this.ExportTo.CheckOnClick = true;
            this.ExportTo.Enabled = false;
            this.ExportTo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExportTo.FormattingEnabled = true;
            this.ExportTo.Location = new System.Drawing.Point(483, 9);
            this.ExportTo.Name = "ExportTo";
            this.ExportTo.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.ExportTo.ScrollAlwaysVisible = true;
            this.ExportTo.Size = new System.Drawing.Size(134, 340);
            this.ExportTo.Sorted = true;
            this.ExportTo.TabIndex = 15;
            this.ExportTo.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.ExportTo_ItemCheck);
            this.ExportTo.SelectedIndexChanged += new System.EventHandler(this.ExportTo_SelectedIndexChanged);
            // 
            // Autofit
            // 
            this.Autofit.AutoSize = true;
            this.Autofit.Location = new System.Drawing.Point(624, 67);
            this.Autofit.Name = "Autofit";
            this.Autofit.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Autofit.Size = new System.Drawing.Size(134, 17);
            this.Autofit.TabIndex = 16;
            this.Autofit.Text = "تنظیم عرض ستونها(کند)";
            this.Autofit.UseVisualStyleBackColor = true;
            // 
            // RemoveDuplicates
            // 
            this.RemoveDuplicates.AutoSize = true;
            this.RemoveDuplicates.Location = new System.Drawing.Point(633, 86);
            this.RemoveDuplicates.Name = "RemoveDuplicates";
            this.RemoveDuplicates.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RemoveDuplicates.Size = new System.Drawing.Size(125, 17);
            this.RemoveDuplicates.TabIndex = 18;
            this.RemoveDuplicates.Text = "تکراری ها حذف شوند";
            this.RemoveDuplicates.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.workBookPassword);
            this.groupBox3.Controls.Add(this.setWorkbookPassword);
            this.groupBox3.Location = new System.Drawing.Point(624, 130);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.groupBox3.Size = new System.Drawing.Size(294, 50);
            this.groupBox3.TabIndex = 19;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "رمز گذاری";
            // 
            // workBookPassword
            // 
            this.workBookPassword.Location = new System.Drawing.Point(9, 19);
            this.workBookPassword.Name = "workBookPassword";
            this.workBookPassword.Size = new System.Drawing.Size(156, 20);
            this.workBookPassword.TabIndex = 1;
            // 
            // setWorkbookPassword
            // 
            this.setWorkbookPassword.AutoSize = true;
            this.setWorkbookPassword.Location = new System.Drawing.Point(197, 22);
            this.setWorkbookPassword.Name = "setWorkbookPassword";
            this.setWorkbookPassword.Size = new System.Drawing.Size(91, 17);
            this.setWorkbookPassword.TabIndex = 0;
            this.setWorkbookPassword.Text = "فعال شدن رمز";
            this.setWorkbookPassword.UseVisualStyleBackColor = true;
            // 
            // freezTopRow
            // 
            this.freezTopRow.AutoSize = true;
            this.freezTopRow.Location = new System.Drawing.Point(640, 107);
            this.freezTopRow.Name = "freezTopRow";
            this.freezTopRow.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.freezTopRow.Size = new System.Drawing.Size(118, 17);
            this.freezTopRow.TabIndex = 20;
            this.freezTopRow.Text = "ثابت کردن ردیف اول";
            this.freezTopRow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.freezTopRow.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(821, 93);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 21;
            this.button1.Text = "Async";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(930, 689);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.freezTopRow);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.RemoveDuplicates);
            this.Controls.Add(this.Autofit);
            this.Controls.Add(this.ExportTo);
            this.Controls.Add(this.logBox);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.Sheet2File);
            this.Controls.Add(this.Sheets2Files);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.Headers);
            this.Controls.Add(this.Worksheets);
            this.Controls.Add(this.btnOpenFile);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainForm";
            this.Text = "Excel Splitter 3";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.ComboBox Worksheets;
        private System.Windows.Forms.ComboBox Headers;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button Sheets2Files;
        private System.Windows.Forms.Button Sheet2File;
        private System.Windows.Forms.CheckBox sendMail;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox MailPart;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox mailBody;
        private System.Windows.Forms.TextBox mailSubject;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ListBox logBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckedListBox ExportTo;
        private System.Windows.Forms.CheckBox Autofit;
        private System.Windows.Forms.CheckBox RemoveDuplicates;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox workBookPassword;
        private System.Windows.Forms.CheckBox setWorkbookPassword;
        private System.Windows.Forms.CheckBox freezTopRow;
        private System.Windows.Forms.Button button1;
    }
}

