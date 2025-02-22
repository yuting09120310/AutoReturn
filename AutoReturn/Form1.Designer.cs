namespace AutoReturn
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
            txtOutput = new TextBox();
            btnStart = new Button();
            txtFilePath = new TextBox();
            SuspendLayout();
            // 
            // txtOutput
            // 
            txtOutput.Location = new Point(31, 96);
            txtOutput.Multiline = true;
            txtOutput.Name = "txtOutput";
            txtOutput.ScrollBars = ScrollBars.Horizontal;
            txtOutput.Size = new Size(460, 465);
            txtOutput.TabIndex = 0;
            // 
            // btnStart
            // 
            btnStart.Location = new Point(412, 44);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(75, 23);
            btnStart.TabIndex = 1;
            btnStart.Text = "開始退貨";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnProcess_Click;
            // 
            // txtFilePath
            // 
            txtFilePath.Location = new Point(31, 44);
            txtFilePath.Name = "txtFilePath";
            txtFilePath.Size = new Size(349, 23);
            txtFilePath.TabIndex = 2;
            txtFilePath.Click += txtFilePath_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(521, 588);
            Controls.Add(txtFilePath);
            Controls.Add(btnStart);
            Controls.Add(txtOutput);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "自動退貨腳本";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txtOutput;
        private Button btnStart;
        private TextBox txtFilePath;
    }
}
