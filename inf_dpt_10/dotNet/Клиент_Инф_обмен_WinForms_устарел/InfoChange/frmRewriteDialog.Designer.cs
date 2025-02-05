namespace InfoChange
{
    partial class frmRewriteDialog
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
            this.btnNo = new System.Windows.Forms.Button();
            this.btnRewrite = new System.Windows.Forms.Button();
            this.btnAppend = new System.Windows.Forms.Button();
            this.cbxAll = new System.Windows.Forms.CheckBox();
            this.lblText = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnNo
            // 
            this.btnNo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnNo.Location = new System.Drawing.Point(10, 41);
            this.btnNo.Name = "btnNo";
            this.btnNo.Size = new System.Drawing.Size(114, 23);
            this.btnNo.TabIndex = 0;
            this.btnNo.Text = "Нет";
            this.btnNo.UseVisualStyleBackColor = true;
            this.btnNo.Click += new System.EventHandler(this.btnNo_Click);
            // 
            // btnRewrite
            // 
            this.btnRewrite.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnRewrite.Location = new System.Drawing.Point(132, 41);
            this.btnRewrite.Name = "btnRewrite";
            this.btnRewrite.Size = new System.Drawing.Size(114, 23);
            this.btnRewrite.TabIndex = 1;
            this.btnRewrite.Text = "Перезаписать";
            this.btnRewrite.UseVisualStyleBackColor = true;
            this.btnRewrite.Click += new System.EventHandler(this.btnRewrite_Click);
            // 
            // btnAppend
            // 
            this.btnAppend.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnAppend.Location = new System.Drawing.Point(253, 41);
            this.btnAppend.Name = "btnAppend";
            this.btnAppend.Size = new System.Drawing.Size(114, 23);
            this.btnAppend.TabIndex = 2;
            this.btnAppend.Text = "Дописать";
            this.btnAppend.UseVisualStyleBackColor = true;
            this.btnAppend.Click += new System.EventHandler(this.btnAppend_Click);
            // 
            // cbxAll
            // 
            this.cbxAll.AutoSize = true;
            this.cbxAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.cbxAll.Location = new System.Drawing.Point(83, 77);
            this.cbxAll.Name = "cbxAll";
            this.cbxAll.Size = new System.Drawing.Size(212, 19);
            this.cbxAll.TabIndex = 3;
            this.cbxAll.Text = "Применить ко всем дальше";
            this.cbxAll.UseVisualStyleBackColor = true;
            this.cbxAll.CheckedChanged += new System.EventHandler(this.cbxAll_CheckedChanged);
            // 
            // lblText
            // 
            this.lblText.AutoSize = true;
            this.lblText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblText.Location = new System.Drawing.Point(7, 9);
            this.lblText.Name = "lblText";
            this.lblText.Size = new System.Drawing.Size(366, 17);
            this.lblText.TabIndex = 4;
            this.lblText.Text = "Внимание! Ответ уже загружен. Перезаписать?";
            this.lblText.Click += new System.EventHandler(this.lblText_Click);
            // 
            // frmRewriteDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(379, 108);
            this.Controls.Add(this.lblText);
            this.Controls.Add(this.cbxAll);
            this.Controls.Add(this.btnAppend);
            this.Controls.Add(this.btnRewrite);
            this.Controls.Add(this.btnNo);
            this.Name = "frmRewriteDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmRewriteDialog";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmRewriteDialog_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnNo;
        private System.Windows.Forms.Button btnRewrite;
        private System.Windows.Forms.Button btnAppend;
        private System.Windows.Forms.CheckBox cbxAll;
        private System.Windows.Forms.Label lblText;
    }
}