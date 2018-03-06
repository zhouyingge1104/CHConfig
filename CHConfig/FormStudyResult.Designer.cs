namespace CHConfig
{
    partial class FormStudyResult
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblStudyResultType = new System.Windows.Forms.Label();
            this.lv = new System.Windows.Forms.ListBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.SystemColors.Control;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.btnClose, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.lblStudyResultType, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.lv, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(1, 1);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.012987F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 92.98701F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 33F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1018, 404);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.SystemColors.Control;
            this.btnClose.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnClose.Font = new System.Drawing.Font("Proxy 9", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(868, 373);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(147, 28);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblStudyResultType
            // 
            this.lblStudyResultType.AutoSize = true;
            this.lblStudyResultType.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblStudyResultType.Font = new System.Drawing.Font("SimSun-ExtB", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStudyResultType.ForeColor = System.Drawing.Color.Black;
            this.lblStudyResultType.Location = new System.Drawing.Point(3, 0);
            this.lblStudyResultType.Name = "lblStudyResultType";
            this.lblStudyResultType.Size = new System.Drawing.Size(1012, 26);
            this.lblStudyResultType.TabIndex = 2;
            this.lblStudyResultType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lv
            // 
            this.lv.BackColor = System.Drawing.Color.White;
            this.lv.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lv.FormattingEnabled = true;
            this.lv.Location = new System.Drawing.Point(3, 29);
            this.lv.Name = "lv";
            this.lv.Size = new System.Drawing.Size(1012, 338);
            this.lv.TabIndex = 3;
            // 
            // FormStudyResult
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.ClientSize = new System.Drawing.Size(1020, 406);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormStudyResult";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FormStudyResult_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btnClose;
        public System.Windows.Forms.Label lblStudyResultType;
        public System.Windows.Forms.ListBox lv;

    }
}