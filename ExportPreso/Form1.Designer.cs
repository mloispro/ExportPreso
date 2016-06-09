namespace ExportPreso
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
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.btnExport = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.lblMessage = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // folderBrowser
            // 
            this.folderBrowser.SelectedPath = "C:\\Projects\\ExportPreso\\Presos";
            this.folderBrowser.ShowNewFolderButton = false;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(12, 34);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 0;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            this.openFileDialog.InitialDirectory = "C:\\Projects\\ExportPreso\\Presos";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 134);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(281, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "name ppts or pdfs with 1_name to order,  -ex:  1_preso.ppt";
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Location = new System.Drawing.Point(9, 77);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(69, 13);
            this.lblMessage.TabIndex = 2;
            this.lblMessage.TabStop = true;
            this.lblMessage.Text = "Select Folder";
            this.lblMessage.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lblMessage_LinkClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(475, 261);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExport);
            this.Name = "Form1";
            this.Text = "Preso Exporter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.LinkLabel lblMessage;
    }
}

