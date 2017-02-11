namespace BayantechAddIn
{
    partial class ApplyFontForm
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
            this.cmb_language = new System.Windows.Forms.ComboBox();
            this.cmb_font_name = new System.Windows.Forms.ComboBox();
            this.lbl_font_name = new System.Windows.Forms.Label();
            this.lbl_apply_on = new System.Windows.Forms.Label();
            this.btn_cancel = new System.Windows.Forms.Button();
            this.btn_apply = new System.Windows.Forms.Button();
            this.cmb_region = new System.Windows.Forms.ComboBox();
            this.lbl_region = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmb_language
            // 
            this.cmb_language.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_language.FormattingEnabled = true;
            this.cmb_language.Location = new System.Drawing.Point(81, 44);
            this.cmb_language.Name = "cmb_language";
            this.cmb_language.Size = new System.Drawing.Size(191, 21);
            this.cmb_language.TabIndex = 1;
            // 
            // cmb_font_name
            // 
            this.cmb_font_name.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_font_name.FormattingEnabled = true;
            this.cmb_font_name.Location = new System.Drawing.Point(81, 17);
            this.cmb_font_name.Name = "cmb_font_name";
            this.cmb_font_name.Size = new System.Drawing.Size(191, 21);
            this.cmb_font_name.TabIndex = 0;
            // 
            // lbl_font_name
            // 
            this.lbl_font_name.AutoSize = true;
            this.lbl_font_name.Location = new System.Drawing.Point(12, 20);
            this.lbl_font_name.Name = "lbl_font_name";
            this.lbl_font_name.Size = new System.Drawing.Size(62, 13);
            this.lbl_font_name.TabIndex = 10;
            this.lbl_font_name.Text = "Font Name:";
            // 
            // lbl_apply_on
            // 
            this.lbl_apply_on.AutoSize = true;
            this.lbl_apply_on.Location = new System.Drawing.Point(12, 47);
            this.lbl_apply_on.Name = "lbl_apply_on";
            this.lbl_apply_on.Size = new System.Drawing.Size(58, 13);
            this.lbl_apply_on.TabIndex = 11;
            this.lbl_apply_on.Text = "Language:";
            // 
            // btn_cancel
            // 
            this.btn_cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btn_cancel.Location = new System.Drawing.Point(197, 103);
            this.btn_cancel.Name = "btn_cancel";
            this.btn_cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_cancel.TabIndex = 3;
            this.btn_cancel.Text = "Cancel";
            this.btn_cancel.UseVisualStyleBackColor = true;
            this.btn_cancel.Click += new System.EventHandler(this.btn_cancel_Click);
            // 
            // btn_apply
            // 
            this.btn_apply.Location = new System.Drawing.Point(81, 103);
            this.btn_apply.Name = "btn_apply";
            this.btn_apply.Size = new System.Drawing.Size(75, 23);
            this.btn_apply.TabIndex = 2;
            this.btn_apply.Text = "Apply";
            this.btn_apply.UseVisualStyleBackColor = true;
            this.btn_apply.Click += new System.EventHandler(this.btn_apply_Click);
            // 
            // cmb_region
            // 
            this.cmb_region.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_region.FormattingEnabled = true;
            this.cmb_region.Location = new System.Drawing.Point(81, 71);
            this.cmb_region.Name = "cmb_region";
            this.cmb_region.Size = new System.Drawing.Size(191, 21);
            this.cmb_region.TabIndex = 12;
            // 
            // lbl_region
            // 
            this.lbl_region.AutoSize = true;
            this.lbl_region.Location = new System.Drawing.Point(12, 74);
            this.lbl_region.Name = "lbl_region";
            this.lbl_region.Size = new System.Drawing.Size(44, 13);
            this.lbl_region.TabIndex = 13;
            this.lbl_region.Text = "Region:";
            // 
            // ApplyFontForm
            // 
            this.AcceptButton = this.btn_apply;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_cancel;
            this.ClientSize = new System.Drawing.Size(284, 135);
            this.ControlBox = false;
            this.Controls.Add(this.cmb_region);
            this.Controls.Add(this.lbl_region);
            this.Controls.Add(this.cmb_language);
            this.Controls.Add(this.cmb_font_name);
            this.Controls.Add(this.lbl_font_name);
            this.Controls.Add(this.lbl_apply_on);
            this.Controls.Add(this.btn_cancel);
            this.Controls.Add(this.btn_apply);
            this.Name = "ApplyFontForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Apply Font";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmb_language;
        private System.Windows.Forms.ComboBox cmb_font_name;
        private System.Windows.Forms.Label lbl_font_name;
        private System.Windows.Forms.Label lbl_apply_on;
        private System.Windows.Forms.Button btn_cancel;
        private System.Windows.Forms.Button btn_apply;
        private System.Windows.Forms.ComboBox cmb_region;
        private System.Windows.Forms.Label lbl_region;
    }
}