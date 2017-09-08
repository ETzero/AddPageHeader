namespace Test
{
    partial class Form1
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
            this.btExcel = new System.Windows.Forms.Button();
            this.btWord = new System.Windows.Forms.Button();
            this.btPPT = new System.Windows.Forms.Button();
            this.tbExcel = new System.Windows.Forms.TextBox();
            this.tbWord = new System.Windows.Forms.TextBox();
            this.tbPPT = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btExcel
            // 
            this.btExcel.Location = new System.Drawing.Point(795, 32);
            this.btExcel.Name = "btExcel";
            this.btExcel.Size = new System.Drawing.Size(144, 39);
            this.btExcel.TabIndex = 0;
            this.btExcel.Text = "Excel加页眉";
            this.btExcel.UseVisualStyleBackColor = true;
            this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
            // 
            // btWord
            // 
            this.btWord.Location = new System.Drawing.Point(795, 93);
            this.btWord.Name = "btWord";
            this.btWord.Size = new System.Drawing.Size(144, 37);
            this.btWord.TabIndex = 1;
            this.btWord.Text = "Word加页眉";
            this.btWord.UseVisualStyleBackColor = true;
            this.btWord.Click += new System.EventHandler(this.btWord_Click);
            // 
            // btPPT
            // 
            this.btPPT.Location = new System.Drawing.Point(795, 149);
            this.btPPT.Name = "btPPT";
            this.btPPT.Size = new System.Drawing.Size(144, 37);
            this.btPPT.TabIndex = 2;
            this.btPPT.Text = "PPT加页眉";
            this.btPPT.UseVisualStyleBackColor = true;
            this.btPPT.Click += new System.EventHandler(this.btPPT_Click);
            // 
            // tbExcel
            // 
            this.tbExcel.Font = new System.Drawing.Font("SimSun", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbExcel.Location = new System.Drawing.Point(30, 36);
            this.tbExcel.Name = "tbExcel";
            this.tbExcel.Size = new System.Drawing.Size(734, 29);
            this.tbExcel.TabIndex = 3;
            // 
            // tbWord
            // 
            this.tbWord.Font = new System.Drawing.Font("SimSun", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbWord.Location = new System.Drawing.Point(30, 101);
            this.tbWord.Name = "tbWord";
            this.tbWord.Size = new System.Drawing.Size(734, 29);
            this.tbWord.TabIndex = 4;
            // 
            // tbPPT
            // 
            this.tbPPT.Font = new System.Drawing.Font("SimSun", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbPPT.Location = new System.Drawing.Point(30, 157);
            this.tbPPT.Name = "tbPPT";
            this.tbPPT.Size = new System.Drawing.Size(734, 29);
            this.tbPPT.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1038, 286);
            this.Controls.Add(this.tbPPT);
            this.Controls.Add(this.tbWord);
            this.Controls.Add(this.tbExcel);
            this.Controls.Add(this.btPPT);
            this.Controls.Add(this.btWord);
            this.Controls.Add(this.btExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btExcel;
        private System.Windows.Forms.Button btWord;
        private System.Windows.Forms.Button btPPT;
        private System.Windows.Forms.TextBox tbExcel;
        private System.Windows.Forms.TextBox tbWord;
        private System.Windows.Forms.TextBox tbPPT;
    }
}

