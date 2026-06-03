namespace utilities
{
    partial class ChangeCaseForm
    {
        private System.Windows.Forms.RadioButton optLowercase;
        private System.Windows.Forms.RadioButton optUppercase;
        private System.Windows.Forms.RadioButton optToggleCase;
        private System.Windows.Forms.RadioButton optSentenceCase;
        private System.Windows.Forms.RadioButton optGlossaryCase;
        private System.Windows.Forms.RadioButton optCapitalizeEachWord;
        private System.Windows.Forms.RadioButton optStartEachWordWithUppercase;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnCancel;

        private void InitializeComponent()
        {
            this.optLowercase = new System.Windows.Forms.RadioButton();
            this.optUppercase = new System.Windows.Forms.RadioButton();
            this.optToggleCase = new System.Windows.Forms.RadioButton();
            this.optSentenceCase = new System.Windows.Forms.RadioButton();
            this.optGlossaryCase = new System.Windows.Forms.RadioButton();
            this.optCapitalizeEachWord = new System.Windows.Forms.RadioButton();
            this.optStartEachWordWithUppercase = new System.Windows.Forms.RadioButton();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();

            // 
            // optLowercase
            // 
            this.optLowercase.Name = "optLowercase";
            this.optLowercase.Text = "Lowercase";
            this.optLowercase.Location = new System.Drawing.Point(20, 20);
            this.optLowercase.AutoSize = true;
            this.optLowercase.TabIndex = 0;

            // 
            // optUppercase
            // 
            this.optUppercase.Name = "optUppercase";
            this.optUppercase.Text = "Uppercase";
            this.optUppercase.Location = new System.Drawing.Point(20, 50);
            this.optUppercase.AutoSize = true;
            this.optUppercase.TabIndex = 1;

            // 
            // optToggleCase
            // 
            this.optToggleCase.Name = "optToggleCase";
            this.optToggleCase.Text = "Toggle Case";
            this.optToggleCase.Location = new System.Drawing.Point(20, 80);
            this.optToggleCase.AutoSize = true;
            this.optToggleCase.TabIndex = 2;

            // 
            // optSentenceCase
            // 
            this.optSentenceCase.Name = "optSentenceCase";
            this.optSentenceCase.Text = "Sentence Case";
            this.optSentenceCase.Location = new System.Drawing.Point(20, 110);
            this.optSentenceCase.AutoSize = true;
            this.optSentenceCase.TabIndex = 3;

            // 
            // optGlossaryCase
            // 
            this.optGlossaryCase.Name = "optGlossaryCase";
            this.optGlossaryCase.Text = "Glossary Case";
            this.optGlossaryCase.Location = new System.Drawing.Point(20, 140);
            this.optGlossaryCase.AutoSize = true;
            this.optGlossaryCase.TabIndex = 4;

            // 
            // optCapitalizeEachWord
            // 
            this.optCapitalizeEachWord.Name = "optCapitalizeEachWord";
            this.optCapitalizeEachWord.Text = "Capitalize Each Word";
            this.optCapitalizeEachWord.Location = new System.Drawing.Point(20, 170);
            this.optCapitalizeEachWord.AutoSize = true;
            this.optCapitalizeEachWord.TabIndex = 5;

            // 
            // optStartEachWordWithUppercase
            // 
            this.optStartEachWordWithUppercase.Name = "optStartEachWordWithUppercase";
            this.optStartEachWordWithUppercase.Text = "Start Each Word With Uppercase";
            this.optStartEachWordWithUppercase.Location = new System.Drawing.Point(20, 200);
            this.optStartEachWordWithUppercase.AutoSize = true;
            this.optStartEachWordWithUppercase.TabIndex = 6;

            // 
            // btnApply
            // 
            this.btnApply.Name = "btnApply";
            this.btnApply.Text = "Apply";
            this.btnApply.Location = new System.Drawing.Point(20, 240);
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);

            // 
            // btnCancel
            // 
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Location = new System.Drawing.Point(120, 240);
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            // 
            // ChangeCaseForm
            // 
            this.Controls.Add(this.optLowercase);
            this.Controls.Add(this.optUppercase);
            this.Controls.Add(this.optToggleCase);
            this.Controls.Add(this.optSentenceCase);
            this.Controls.Add(this.optGlossaryCase);
            this.Controls.Add(this.optCapitalizeEachWord);
            this.Controls.Add(this.optStartEachWordWithUppercase);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Name = "ChangeCaseForm";
            this.Text = "Change Case";
            this.ClientSize = new System.Drawing.Size(300, 300);
        }
    }
}
