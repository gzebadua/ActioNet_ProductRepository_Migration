namespace ProductRepository_Migration
{
    partial class CreateFlatKeywordsTable
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
            this.btnFlat = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnFlat
            // 
            this.btnFlat.Location = new System.Drawing.Point(80, 62);
            this.btnFlat.Name = "btnFlat";
            this.btnFlat.Size = new System.Drawing.Size(75, 23);
            this.btnFlat.TabIndex = 0;
            this.btnFlat.Text = "Flatten";
            this.btnFlat.UseVisualStyleBackColor = true;
            this.btnFlat.Click += new System.EventHandler(this.btnFlat_Click);
            // 
            // CreateFlatKeywordsTable
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 262);
            this.Controls.Add(this.btnFlat);
            this.Name = "CreateFlatKeywordsTable";
            this.Text = "CreateFlatKeywordsTable";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnFlat;
    }
}