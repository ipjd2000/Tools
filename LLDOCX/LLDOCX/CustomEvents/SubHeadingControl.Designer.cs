
namespace LLDOCX.CustomEvents
{
    partial class SubHeadingControl
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
            this.label1 = new System.Windows.Forms.Label();
            this.comboxSubHeading = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Left;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "label1";
            // 
            // comboxSubHeading
            // 
            this.comboxSubHeading.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboxSubHeading.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboxSubHeading.FormattingEnabled = true;
            this.comboxSubHeading.Location = new System.Drawing.Point(55, 0);
            this.comboxSubHeading.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.comboxSubHeading.Name = "comboxSubHeading";
            this.comboxSubHeading.Size = new System.Drawing.Size(396, 31);
            this.comboxSubHeading.TabIndex = 1;
            this.comboxSubHeading.SelectionChangeCommitted += new System.EventHandler(this.OncomboxSubHeadingSelectionChangeCommitted);
            // 
            // SubHeadingControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 23F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.comboxSubHeading);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SubHeadingControl";
            this.Size = new System.Drawing.Size(451, 41);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboxSubHeading;
    }
}
