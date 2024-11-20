namespace XpandNPIManager
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Button buttonListFiles;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.buttonListFiles = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();

            // 
            // buttonListFiles
            // 
            this.buttonListFiles.Location = new System.Drawing.Point(100, 50);
            this.buttonListFiles.Name = "buttonListFiles";
            this.buttonListFiles.Size = new System.Drawing.Size(150, 50);
            this.buttonListFiles.TabIndex = 0;
            this.buttonListFiles.Text = "List All Files";
            this.buttonListFiles.UseVisualStyleBackColor = true;
            this.buttonListFiles.Click += new System.EventHandler(this.buttonListFiles_Click);

            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(50, 120);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(250, 30);
            this.progressBar.TabIndex = 1;

            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(50, 160);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(100, 17);
            this.lblStatus.TabIndex = 2;
            this.lblStatus.Text = "Status: Ready";

            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(350, 250);
            this.Controls.Add(this.buttonListFiles);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.Name = "Form1";
            this.Text = "Xpand NPI Manager";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
