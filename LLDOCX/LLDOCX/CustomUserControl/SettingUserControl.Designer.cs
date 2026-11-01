namespace LLDOCX.CustomUserControl
{
    partial class SettingUserControl
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.ckbPageSize = new System.Windows.Forms.CheckBox();
            this.ckbPageOrientation = new System.Windows.Forms.CheckBox();
            this.ckbMarkup = new System.Windows.Forms.CheckBox();
            this.ckbhyperlink = new System.Windows.Forms.CheckBox();
            this.ckbComment = new System.Windows.Forms.CheckBox();
            this.ckbBookmarks = new System.Windows.Forms.CheckBox();
            this.ckbDocProperty = new System.Windows.Forms.CheckBox();
            this.ckbKeepTracking = new System.Windows.Forms.CheckBox();
            this.cbkHeadingValidation = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.ckbCheckAll = new System.Windows.Forms.CheckBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.ckbFont = new System.Windows.Forms.CheckBox();
            this.btnFontName = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5F));
            this.tableLayoutPanel1.Controls.Add(this.ckbPageSize, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.ckbPageOrientation, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.ckbMarkup, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.ckbhyperlink, 1, 7);
            this.tableLayoutPanel1.Controls.Add(this.ckbComment, 1, 8);
            this.tableLayoutPanel1.Controls.Add(this.ckbBookmarks, 1, 9);
            this.tableLayoutPanel1.Controls.Add(this.ckbDocProperty, 1, 10);
            this.tableLayoutPanel1.Controls.Add(this.ckbKeepTracking, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.cbkHeadingValidation, 1, 11);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 1, 13);
            this.tableLayoutPanel1.Controls.Add(this.ckbCheckAll, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 1, 6);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 14;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 8F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 51F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 14F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(322, 554);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // ckbPageSize
            // 
            this.ckbPageSize.AutoSize = true;
            this.ckbPageSize.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbPageSize.Location = new System.Drawing.Point(8, 44);
            this.ckbPageSize.Name = "ckbPageSize";
            this.ckbPageSize.Size = new System.Drawing.Size(306, 27);
            this.ckbPageSize.TabIndex = 0;
            this.ckbPageSize.Text = "Check Page Size (A4)";
            this.ckbPageSize.UseVisualStyleBackColor = true;
            // 
            // ckbPageOrientation
            // 
            this.ckbPageOrientation.AutoSize = true;
            this.ckbPageOrientation.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbPageOrientation.Location = new System.Drawing.Point(8, 77);
            this.ckbPageOrientation.Name = "ckbPageOrientation";
            this.ckbPageOrientation.Size = new System.Drawing.Size(306, 27);
            this.ckbPageOrientation.TabIndex = 1;
            this.ckbPageOrientation.Text = "Check Page Orientation (Portrait)";
            this.ckbPageOrientation.UseVisualStyleBackColor = true;
            // 
            // ckbMarkup
            // 
            this.ckbMarkup.AutoSize = true;
            this.ckbMarkup.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbMarkup.Location = new System.Drawing.Point(8, 110);
            this.ckbMarkup.Name = "ckbMarkup";
            this.ckbMarkup.Size = new System.Drawing.Size(306, 27);
            this.ckbMarkup.TabIndex = 2;
            this.ckbMarkup.Text = "Check Mark-ups (Revisions)";
            this.ckbMarkup.UseVisualStyleBackColor = true;
            // 
            // ckbhyperlink
            // 
            this.ckbhyperlink.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbhyperlink.Location = new System.Drawing.Point(8, 222);
            this.ckbhyperlink.Name = "ckbhyperlink";
            this.ckbhyperlink.Size = new System.Drawing.Size(306, 24);
            this.ckbhyperlink.TabIndex = 4;
            this.ckbhyperlink.Text = "Check for Hyperlinks, AutoNumbers, && Fields";
            this.ckbhyperlink.UseVisualStyleBackColor = true;
            // 
            // ckbComment
            // 
            this.ckbComment.AutoSize = true;
            this.ckbComment.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbComment.Location = new System.Drawing.Point(8, 252);
            this.ckbComment.Name = "ckbComment";
            this.ckbComment.Size = new System.Drawing.Size(306, 27);
            this.ckbComment.TabIndex = 5;
            this.ckbComment.Text = "Check for Comments";
            this.ckbComment.UseVisualStyleBackColor = true;
            // 
            // ckbBookmarks
            // 
            this.ckbBookmarks.AutoSize = true;
            this.ckbBookmarks.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbBookmarks.Location = new System.Drawing.Point(8, 285);
            this.ckbBookmarks.Name = "ckbBookmarks";
            this.ckbBookmarks.Size = new System.Drawing.Size(306, 27);
            this.ckbBookmarks.TabIndex = 6;
            this.ckbBookmarks.Text = "Check for Bookmarks";
            this.ckbBookmarks.UseVisualStyleBackColor = true;
            // 
            // ckbDocProperty
            // 
            this.ckbDocProperty.AutoSize = true;
            this.ckbDocProperty.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbDocProperty.Location = new System.Drawing.Point(8, 318);
            this.ckbDocProperty.Name = "ckbDocProperty";
            this.ckbDocProperty.Size = new System.Drawing.Size(306, 27);
            this.ckbDocProperty.TabIndex = 7;
            this.ckbDocProperty.Text = "Remove document Properties";
            this.ckbDocProperty.UseVisualStyleBackColor = true;
            // 
            // ckbKeepTracking
            // 
            this.ckbKeepTracking.AutoSize = true;
            this.ckbKeepTracking.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbKeepTracking.Location = new System.Drawing.Point(25, 143);
            this.ckbKeepTracking.Margin = new System.Windows.Forms.Padding(20, 3, 3, 3);
            this.ckbKeepTracking.Name = "ckbKeepTracking";
            this.ckbKeepTracking.Size = new System.Drawing.Size(289, 27);
            this.ckbKeepTracking.TabIndex = 8;
            this.ckbKeepTracking.Text = "Keep Tracking ON";
            this.ckbKeepTracking.UseVisualStyleBackColor = true;
            // 
            // cbkHeadingValidation
            // 
            this.cbkHeadingValidation.AutoSize = true;
            this.cbkHeadingValidation.Location = new System.Drawing.Point(8, 351);
            this.cbkHeadingValidation.Name = "cbkHeadingValidation";
            this.cbkHeadingValidation.Size = new System.Drawing.Size(177, 27);
            this.cbkHeadingValidation.TabIndex = 9;
            this.cbkHeadingValidation.Text = "Heading Validation";
            this.cbkHeadingValidation.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(8, 435);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(306, 116);
            this.panel1.TabIndex = 10;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.btnSave);
            this.flowLayoutPanel1.Controls.Add(this.btnReset);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 505);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(322, 49);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(202, 3);
            this.btnSave.Margin = new System.Windows.Forms.Padding(3, 3, 10, 3);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(110, 40);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.OnBtnSaveChangeClicked);
            // 
            // btnReset
            // 
            this.btnReset.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReset.Location = new System.Drawing.Point(74, 3);
            this.btnReset.Margin = new System.Windows.Forms.Padding(3, 3, 15, 3);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(110, 40);
            this.btnReset.TabIndex = 1;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.OnButnResetClicked);
            // 
            // ckbCheckAll
            // 
            this.ckbCheckAll.AutoSize = true;
            this.ckbCheckAll.Location = new System.Drawing.Point(8, 11);
            this.ckbCheckAll.Name = "ckbCheckAll";
            this.ckbCheckAll.Size = new System.Drawing.Size(120, 27);
            this.ckbCheckAll.TabIndex = 11;
            this.ckbCheckAll.Text = "Selected All";
            this.ckbCheckAll.UseVisualStyleBackColor = true;
            this.ckbCheckAll.CheckedChanged += new System.EventHandler(this.ckbCheckAll_CheckedChanged);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnFontName);
            this.panel2.Controls.Add(this.ckbFont);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(8, 176);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(306, 40);
            this.panel2.TabIndex = 12;
            // 
            // ckbFont
            // 
            this.ckbFont.AutoSize = true;
            this.ckbFont.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ckbFont.Location = new System.Drawing.Point(0, 0);
            this.ckbFont.Name = "ckbFont";
            this.ckbFont.Size = new System.Drawing.Size(306, 40);
            this.ckbFont.TabIndex = 4;
            this.ckbFont.Text = "Check Font && Color";
            this.ckbFont.UseVisualStyleBackColor = true;
            this.ckbFont.CheckedChanged += new System.EventHandler(this.ckbFont_CheckedChanged);
            // 
            // btnFontName
            // 
            this.btnFontName.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnFontName.Enabled = false;
            this.btnFontName.Location = new System.Drawing.Point(197, 0);
            this.btnFontName.Name = "btnFontName";
            this.btnFontName.Size = new System.Drawing.Size(109, 40);
            this.btnFontName.TabIndex = 5;
            this.btnFontName.Text = "Select Font";
            this.btnFontName.UseVisualStyleBackColor = true;
            this.btnFontName.Click += new System.EventHandler(this.btnFontName_Click);
            // 
            // SettingUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "SettingUserControl";
            this.Size = new System.Drawing.Size(322, 554);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.CheckBox ckbPageSize;
        private System.Windows.Forms.CheckBox ckbPageOrientation;
        private System.Windows.Forms.CheckBox ckbMarkup;
        private System.Windows.Forms.CheckBox ckbhyperlink;
        private System.Windows.Forms.CheckBox ckbComment;
        private System.Windows.Forms.CheckBox ckbBookmarks;
        private System.Windows.Forms.CheckBox ckbDocProperty;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.CheckBox ckbKeepTracking;
        private System.Windows.Forms.CheckBox cbkHeadingValidation;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox ckbCheckAll;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckBox ckbFont;
        private System.Windows.Forms.Button btnFontName;
    }
}
