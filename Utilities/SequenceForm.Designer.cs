namespace utilities
{

    partial class SequenceForm
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
            this.comboBoxFillOrder = new System.Windows.Forms.ComboBox();
            this.txtStartNumber = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtIncrement = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtDigits = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSuffix = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtPrefix = new System.Windows.Forms.TextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnGeneratePreview = new System.Windows.Forms.Button();
            this.listBoxPreview = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // comboBoxFillOrder
            // 
            this.comboBoxFillOrder.FormattingEnabled = true;
            this.comboBoxFillOrder.Items.AddRange(new object[] {
            "Fill Vertically",
            "Fill Horizontally"});
            this.comboBoxFillOrder.Location = new System.Drawing.Point(23, 22);
            this.comboBoxFillOrder.Name = "comboBoxFillOrder";
            this.comboBoxFillOrder.Size = new System.Drawing.Size(121, 21);
            this.comboBoxFillOrder.TabIndex = 0;
            this.comboBoxFillOrder.Text = "Fill Order";
            // 
            // txtStartNumber
            // 
            this.txtStartNumber.Location = new System.Drawing.Point(298, 23);
            this.txtStartNumber.Name = "txtStartNumber";
            this.txtStartNumber.Size = new System.Drawing.Size(100, 20);
            this.txtStartNumber.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(207, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Start Number";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(207, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Increment";
            // 
            // txtIncrement
            // 
            this.txtIncrement.Location = new System.Drawing.Point(298, 60);
            this.txtIncrement.Name = "txtIncrement";
            this.txtIncrement.Size = new System.Drawing.Size(100, 20);
            this.txtIncrement.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(207, 102);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Number of Digits";
            // 
            // txtDigits
            // 
            this.txtDigits.Location = new System.Drawing.Point(298, 95);
            this.txtDigits.Name = "txtDigits";
            this.txtDigits.Size = new System.Drawing.Size(100, 20);
            this.txtDigits.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(259, 178);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(33, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Suffix";
            // 
            // txtSuffix
            // 
            this.txtSuffix.Location = new System.Drawing.Point(298, 171);
            this.txtSuffix.Name = "txtSuffix";
            this.txtSuffix.Size = new System.Drawing.Size(100, 20);
            this.txtSuffix.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(259, 139);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(33, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Prefix";
            // 
            // txtPrefix
            // 
            this.txtPrefix.Location = new System.Drawing.Point(298, 132);
            this.txtPrefix.Name = "txtPrefix";
            this.txtPrefix.Size = new System.Drawing.Size(100, 20);
            this.txtPrefix.TabIndex = 9;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(685, 45);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(685, 134);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnGeneratePreview
            // 
            this.btnGeneratePreview.Location = new System.Drawing.Point(685, 92);
            this.btnGeneratePreview.Name = "btnGeneratePreview";
            this.btnGeneratePreview.Size = new System.Drawing.Size(75, 23);
            this.btnGeneratePreview.TabIndex = 13;
            this.btnGeneratePreview.Text = "Preview";
            this.btnGeneratePreview.UseVisualStyleBackColor = true;
            // 
            // listBoxPreview
            // 
            this.listBoxPreview.FormattingEnabled = true;
            this.listBoxPreview.Location = new System.Drawing.Point(463, 30);
            this.listBoxPreview.Name = "listBoxPreview";
            this.listBoxPreview.Size = new System.Drawing.Size(199, 329);
            this.listBoxPreview.TabIndex = 14;
            // 
            // SequenceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.listBoxPreview);
            this.Controls.Add(this.btnGeneratePreview);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtPrefix);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtSuffix);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtDigits);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtIncrement);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtStartNumber);
            this.Controls.Add(this.comboBoxFillOrder);
            this.Name = "SequenceForm";
            this.Text = "SequenceForm";
            this.Load += new System.EventHandler(this.SequenceForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxFillOrder;
        private System.Windows.Forms.TextBox txtStartNumber;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtIncrement;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtDigits;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSuffix;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtPrefix;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnGeneratePreview;
        private System.Windows.Forms.ListBox listBoxPreview;
    }
}
