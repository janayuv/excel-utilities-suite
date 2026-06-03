using System;

namespace utilities
{
    partial class TrimTextForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.RadioButton optDeleteLeadingTrailingSpaces;
        private System.Windows.Forms.RadioButton optDeleteLeadingTrailingExcessiveSpaces;
        private System.Windows.Forms.RadioButton optDeleteLeadingCharacters;
        private System.Windows.Forms.RadioButton optDeleteEndingCharacters;
        private System.Windows.Forms.TextBox txtNumberOfCharacters;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnCancel;

        private void optDeleteLeadingCharacters_CheckedChanged(object sender, EventArgs e)
        {
            // Show the textbox only if the "Delete Leading Characters" option is selected
            txtNumberOfCharacters.Visible = optDeleteLeadingCharacters.Checked || optDeleteEndingCharacters.Checked;
        }

        private void optDeleteEndingCharacters_CheckedChanged(object sender, EventArgs e)
        {
            // Show the textbox only if the "Delete Ending Characters" option is selected
            txtNumberOfCharacters.Visible = optDeleteLeadingCharacters.Checked || optDeleteEndingCharacters.Checked;
        }

        private void InitializeComponent()
        {
            this.optDeleteLeadingTrailingSpaces = new System.Windows.Forms.RadioButton();
            this.optDeleteLeadingTrailingExcessiveSpaces = new System.Windows.Forms.RadioButton();
            this.optDeleteLeadingCharacters = new System.Windows.Forms.RadioButton();
            this.optDeleteEndingCharacters = new System.Windows.Forms.RadioButton();
            this.txtNumberOfCharacters = new System.Windows.Forms.TextBox();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            
            // 
            // optDeleteLeadingTrailingSpaces
            // 
            this.optDeleteLeadingTrailingSpaces.AutoSize = true;
            this.optDeleteLeadingTrailingSpaces.Location = new System.Drawing.Point(20, 20);
            this.optDeleteLeadingTrailingSpaces.Name = "optDeleteLeadingTrailingSpaces";
            this.optDeleteLeadingTrailingSpaces.Size = new System.Drawing.Size(196, 17);
            this.optDeleteLeadingTrailingSpaces.TabIndex = 0;
            this.optDeleteLeadingTrailingSpaces.TabStop = true;
            this.optDeleteLeadingTrailingSpaces.Text = "Delete Leading/Trailing Spaces";
            this.optDeleteLeadingTrailingSpaces.UseVisualStyleBackColor = true;
            // 
            // optDeleteLeadingTrailingExcessiveSpaces
            // 
            this.optDeleteLeadingTrailingExcessiveSpaces.AutoSize = true;
            this.optDeleteLeadingTrailingExcessiveSpaces.Location = new System.Drawing.Point(20, 50);
            this.optDeleteLeadingTrailingExcessiveSpaces.Name = "optDeleteLeadingTrailingExcessiveSpaces";
            this.optDeleteLeadingTrailingExcessiveSpaces.Size = new System.Drawing.Size(267, 17);
            this.optDeleteLeadingTrailingExcessiveSpaces.TabIndex = 1;
            this.optDeleteLeadingTrailingExcessiveSpaces.TabStop = true;
            this.optDeleteLeadingTrailingExcessiveSpaces.Text = "Delete Leading/Trailing & Excessive Spaces";
            this.optDeleteLeadingTrailingExcessiveSpaces.UseVisualStyleBackColor = true;
            // 
            // optDeleteLeadingCharacters
            // 
            this.optDeleteLeadingCharacters.AutoSize = true;
            this.optDeleteLeadingCharacters.Location = new System.Drawing.Point(20, 80);
            this.optDeleteLeadingCharacters.Name = "optDeleteLeadingCharacters";
            this.optDeleteLeadingCharacters.Size = new System.Drawing.Size(181, 17);
            this.optDeleteLeadingCharacters.TabIndex = 2;
            this.optDeleteLeadingCharacters.TabStop = true;
            this.optDeleteLeadingCharacters.Text = "Delete Leading Characters";
            this.optDeleteLeadingCharacters.UseVisualStyleBackColor = true;
            this.optDeleteLeadingCharacters.CheckedChanged += new System.EventHandler(this.optDeleteLeadingCharacters_CheckedChanged);
            // 
            // optDeleteEndingCharacters
            // 
            this.optDeleteEndingCharacters.AutoSize = true;
            this.optDeleteEndingCharacters.Location = new System.Drawing.Point(20, 110);
            this.optDeleteEndingCharacters.Name = "optDeleteEndingCharacters";
            this.optDeleteEndingCharacters.Size = new System.Drawing.Size(177, 17);
            this.optDeleteEndingCharacters.TabIndex = 3;
            this.optDeleteEndingCharacters.TabStop = true;
            this.optDeleteEndingCharacters.Text = "Delete Ending Characters";
            this.optDeleteEndingCharacters.UseVisualStyleBackColor = true;
            this.optDeleteEndingCharacters.CheckedChanged += new System.EventHandler(this.optDeleteEndingCharacters_CheckedChanged);
            // 
            // txtNumberOfCharacters
            // 
            this.txtNumberOfCharacters.Location = new System.Drawing.Point(20, 140);
            this.txtNumberOfCharacters.Name = "txtNumberOfCharacters";
            this.txtNumberOfCharacters.Size = new System.Drawing.Size(100, 20);
            this.txtNumberOfCharacters.TabIndex = 4;
            this.txtNumberOfCharacters.Visible = false;
            // 
            // btnApply
            // 
            this.btnApply.Location = new System.Drawing.Point(20, 180);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 23);
            this.btnApply.TabIndex = 5;
            this.btnApply.Text = "Apply";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(120, 180);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // TrimTextForm
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.txtNumberOfCharacters);
            this.Controls.Add(this.optDeleteEndingCharacters);
            this.Controls.Add(this.optDeleteLeadingCharacters);
            this.Controls.Add(this.optDeleteLeadingTrailingExcessiveSpaces);
            this.Controls.Add(this.optDeleteLeadingTrailingSpaces);
            this.Name = "TrimTextForm";
            this.Text = "Trim Text Options";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
