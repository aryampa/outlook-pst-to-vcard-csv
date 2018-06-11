namespace Outlook_TO_Vcf_CSV_EXCEL
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
            this.BtnLoad = new System.Windows.Forms.Button();
            this.progBar = new System.Windows.Forms.ProgressBar();
            this.lbl = new System.Windows.Forms.Label();
            this.bgWorker = new System.ComponentModel.BackgroundWorker();
            this.folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.lbx1 = new System.Windows.Forms.ListBox();
            this.rbtnVcf = new System.Windows.Forms.RadioButton();
            this.RbtnCsv = new System.Windows.Forms.RadioButton();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.lblProgressText = new System.Windows.Forms.Label();
            this.tbxFileName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BtnLoad
            // 
            this.BtnLoad.Location = new System.Drawing.Point(328, 55);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(165, 40);
            this.BtnLoad.TabIndex = 0;
            this.BtnLoad.Text = "Generate Vcf File";
            this.BtnLoad.UseVisualStyleBackColor = true;
            this.BtnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // progBar
            // 
            this.progBar.Location = new System.Drawing.Point(13, 13);
            this.progBar.Name = "progBar";
            this.progBar.Size = new System.Drawing.Size(484, 23);
            this.progBar.TabIndex = 2;
            // 
            // lbl
            // 
            this.lbl.AutoSize = true;
            this.lbl.Location = new System.Drawing.Point(207, 39);
            this.lbl.Name = "lbl";
            this.lbl.Size = new System.Drawing.Size(62, 13);
            this.lbl.TabIndex = 3;
            this.lbl.Text = "xxxxxxxxxxx";
            // 
            // bgWorker
            // 
            this.bgWorker.WorkerReportsProgress = true;
            // 
            // lbx1
            // 
            this.lbx1.FormattingEnabled = true;
            this.lbx1.Location = new System.Drawing.Point(26, 182);
            this.lbx1.Name = "lbx1";
            this.lbx1.ScrollAlwaysVisible = true;
            this.lbx1.Size = new System.Drawing.Size(471, 56);
            this.lbx1.TabIndex = 4;
            // 
            // rbtnVcf
            // 
            this.rbtnVcf.AutoSize = true;
            this.rbtnVcf.Location = new System.Drawing.Point(13, 43);
            this.rbtnVcf.Name = "rbtnVcf";
            this.rbtnVcf.Size = new System.Drawing.Size(60, 17);
            this.rbtnVcf.TabIndex = 5;
            this.rbtnVcf.TabStop = true;
            this.rbtnVcf.Text = "Vcf File";
            this.rbtnVcf.UseVisualStyleBackColor = true;
            // 
            // RbtnCsv
            // 
            this.RbtnCsv.AutoSize = true;
            this.RbtnCsv.Location = new System.Drawing.Point(13, 67);
            this.RbtnCsv.Name = "RbtnCsv";
            this.RbtnCsv.Size = new System.Drawing.Size(46, 17);
            this.RbtnCsv.TabIndex = 6;
            this.RbtnCsv.TabStop = true;
            this.RbtnCsv.Text = "CSV";
            this.RbtnCsv.UseVisualStyleBackColor = true;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(13, 91);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(59, 17);
            this.radioButton3.TabIndex = 7;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "EXCEL";
            this.radioButton3.UseVisualStyleBackColor = true;
            // 
            // lblProgressText
            // 
            this.lblProgressText.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.lblProgressText.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProgressText.Location = new System.Drawing.Point(226, 148);
            this.lblProgressText.Name = "lblProgressText";
            this.lblProgressText.Size = new System.Drawing.Size(267, 23);
            this.lblProgressText.TabIndex = 9;
            this.lblProgressText.Text = "--------";
            this.lblProgressText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tbxFileName
            // 
            this.tbxFileName.Location = new System.Drawing.Point(26, 156);
            this.tbxFileName.Name = "tbxFileName";
            this.tbxFileName.Size = new System.Drawing.Size(168, 20);
            this.tbxFileName.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 137);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "File Name:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(509, 249);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbxFileName);
            this.Controls.Add(this.lblProgressText);
            this.Controls.Add(this.radioButton3);
            this.Controls.Add(this.RbtnCsv);
            this.Controls.Add(this.rbtnVcf);
            this.Controls.Add(this.lbx1);
            this.Controls.Add(this.lbl);
            this.Controls.Add(this.progBar);
            this.Controls.Add(this.BtnLoad);
            this.Name = "Form1";
            this.Text = "OutLook To Vcf";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnLoad;
        private System.Windows.Forms.ProgressBar progBar;
        private System.Windows.Forms.Label lbl;
        private System.ComponentModel.BackgroundWorker bgWorker;
        private System.Windows.Forms.FolderBrowserDialog folderDialog;
        private System.Windows.Forms.ListBox lbx1;
        private System.Windows.Forms.RadioButton rbtnVcf;
        private System.Windows.Forms.RadioButton RbtnCsv;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.Label lblProgressText;
        private System.Windows.Forms.TextBox tbxFileName;
        private System.Windows.Forms.Label label1;
    }
}

