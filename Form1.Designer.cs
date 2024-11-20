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
            this.btnGetFileLocations = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonListFiles
            // 
            this.buttonListFiles.Location = new System.Drawing.Point(112, 62);
            this.buttonListFiles.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.buttonListFiles.Name = "buttonListFiles";
            this.buttonListFiles.Size = new System.Drawing.Size(169, 62);
            this.buttonListFiles.TabIndex = 0;
            this.buttonListFiles.Text = "List All Files";
            this.buttonListFiles.UseVisualStyleBackColor = true;
            this.buttonListFiles.Click += new System.EventHandler(this.buttonListFiles_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(56, 150);
            this.progressBar.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(281, 38);
            this.progressBar.TabIndex = 1;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(56, 200);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(110, 20);
            this.lblStatus.TabIndex = 2;
            this.lblStatus.Text = "Status: Ready";
            // 
            // btnGetFileLocations
            // 
            this.btnGetFileLocations.Location = new System.Drawing.Point(184, 373);
            this.btnGetFileLocations.Name = "btnGetFileLocations";
            this.btnGetFileLocations.Size = new System.Drawing.Size(203, 122);
            this.btnGetFileLocations.TabIndex = 3;
            this.btnGetFileLocations.Text = "Get File Locations";
            this.btnGetFileLocations.UseVisualStyleBackColor = true;
            this.btnGetFileLocations.Click += new System.EventHandler(this.btnGetFileLocations_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1018, 1034);
            this.Controls.Add(this.btnGetFileLocations);
            this.Controls.Add(this.buttonListFiles);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Form1";
            this.Text = "Xpand NPI Manager";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Button btnGetFileLocations;
    }
}
