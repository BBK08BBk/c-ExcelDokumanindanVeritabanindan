namespace ExcelVtEntegrasyonProje
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnvtdenoku = new Button();
            richTextBox1 = new RichTextBox();
            richTextBox2 = new RichTextBox();
            btnexceloku = new Button();
            SuspendLayout();
            // 
            // btnvtdenoku
            // 
            btnvtdenoku.Location = new Point(602, 122);
            btnvtdenoku.Name = "btnvtdenoku";
            btnvtdenoku.Size = new Size(80, 36);
            btnvtdenoku.TabIndex = 0;
            btnvtdenoku.Text = "Verileri çek";
            btnvtdenoku.UseVisualStyleBackColor = true;
            btnvtdenoku.Click += btnvtdenoku_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(155, 99);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(416, 96);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(155, 243);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(416, 96);
            richTextBox2.TabIndex = 2;
            richTextBox2.Text = "";
            // 
            // btnexceloku
            // 
            btnexceloku.Location = new Point(602, 278);
            btnexceloku.Name = "btnexceloku";
            btnexceloku.Size = new Size(116, 38);
            btnexceloku.TabIndex = 3;
            btnexceloku.Text = "Verileri Veritabnına yazdır";
            btnexceloku.UseVisualStyleBackColor = true;
            btnexceloku.Click += btnexceloku_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.DarkGreen;
            ClientSize = new Size(837, 471);
            Controls.Add(btnexceloku);
            Controls.Add(richTextBox2);
            Controls.Add(richTextBox1);
            Controls.Add(btnvtdenoku);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Excelden Veritabanı işlemleri";
            ResumeLayout(false);
        }

        #endregion

        private Button btnvtdenoku;
        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private Button btnexceloku;
    }
}