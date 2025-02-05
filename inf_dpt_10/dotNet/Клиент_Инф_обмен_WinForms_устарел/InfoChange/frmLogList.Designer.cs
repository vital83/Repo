namespace InfoChange
{
    partial class frmLogList
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
            this.dgvwLogList = new System.Windows.Forms.DataGridView();
            this.btnSelect = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.lblTitle1 = new System.Windows.Forms.Label();
            this.lblTitleAgr = new System.Windows.Forms.Label();
            this.btnContinueLoad = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvwLogList)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvwLogList
            // 
            this.dgvwLogList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvwLogList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvwLogList.Location = new System.Drawing.Point(12, 79);
            this.dgvwLogList.MultiSelect = false;
            this.dgvwLogList.Name = "dgvwLogList";
            this.dgvwLogList.Size = new System.Drawing.Size(435, 196);
            this.dgvwLogList.TabIndex = 0;
            this.dgvwLogList.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvwLogList_CellDoubleClick);
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(371, 290);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(75, 23);
            this.btnSelect.TabIndex = 1;
            this.btnSelect.Text = "Выбрать";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(12, 290);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(169, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Отменить загрузку ответа";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblTitle.Location = new System.Drawing.Point(12, 21);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(0, 13);
            this.lblTitle.TabIndex = 3;
            // 
            // lblTitle1
            // 
            this.lblTitle1.AutoSize = true;
            this.lblTitle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblTitle1.Location = new System.Drawing.Point(12, 2);
            this.lblTitle1.Name = "lblTitle1";
            this.lblTitle1.Size = new System.Drawing.Size(184, 13);
            this.lblTitle1.TabIndex = 4;
            this.lblTitle1.Text = "Список неотвеченных запросов из";
            // 
            // lblTitleAgr
            // 
            this.lblTitleAgr.AutoSize = true;
            this.lblTitleAgr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblTitleAgr.Location = new System.Drawing.Point(12, 50);
            this.lblTitleAgr.Name = "lblTitleAgr";
            this.lblTitleAgr.Size = new System.Drawing.Size(0, 13);
            this.lblTitleAgr.TabIndex = 5;
            // 
            // btnContinueLoad
            // 
            this.btnContinueLoad.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnContinueLoad.Location = new System.Drawing.Point(187, 290);
            this.btnContinueLoad.Name = "btnContinueLoad";
            this.btnContinueLoad.Size = new System.Drawing.Size(178, 23);
            this.btnContinueLoad.TabIndex = 3;
            this.btnContinueLoad.Text = "Загрузить ответ без привязки";
            this.btnContinueLoad.UseVisualStyleBackColor = true;
            this.btnContinueLoad.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmLogList
            // 
            this.AcceptButton = this.btnSelect;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(459, 325);
            this.Controls.Add(this.btnContinueLoad);
            this.Controls.Add(this.lblTitleAgr);
            this.Controls.Add(this.lblTitle1);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.dgvwLogList);
            this.Name = "frmLogList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Выбор запроса";
            this.Load += new System.EventHandler(this.frmLogList_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvwLogList)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvwLogList;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblTitle1;
        private System.Windows.Forms.Label lblTitleAgr;
        private System.Windows.Forms.Button btnContinueLoad;
    }
}