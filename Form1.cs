using System;
using System.Windows.Forms;
using NPILib;
using System.Collections.Generic;
using System.IO;

namespace XpandNPIManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreateZip_Click(object sender, EventArgs e)
        {
            try
            {
                // Step 1: Prompt the user to paste a list of part numbers
                string partNumbersInput = PromptForPartNumbers();
                if (string.IsNullOrWhiteSpace(partNumbersInput))
                {
                    MessageBox.Show("No part numbers provided. Operation canceled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                List<string> partNumbers = new List<string>(partNumbersInput.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries));

                // Step 2: Open a folder browser dialog to select the output ZIP location
                string outputZipPath = PromptForOutputZipLocation();
                if (string.IsNullOrWhiteSpace(outputZipPath))
                {
                    MessageBox.Show("No output location selected. Operation canceled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Step 3: Scan \\PDM\ERP-Data\Files to create "Files"
                string scanPath = @"\\PDM\ERP-Data\Files";
                if (!Directory.Exists(scanPath))
                {
                    MessageBox.Show($"The directory '{scanPath}' does not exist. Operation canceled.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                List<ProdFile> files = ScanFiles(scanPath);

                // Step 4: Create PNFiles using the part numbers and files
                List<PNFiles> pnFilesList = PNFileProcessor.GetPNFilesForPartNumbers(files, partNumbers);

                // Step 5: Create the ZIP file
                PNFileCompressor.CompressFiles(pnFilesList, outputZipPath);

                MessageBox.Show($"ZIP file created successfully at: {outputZipPath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private string PromptForPartNumbers()
        {
            using (Form inputForm = new Form())
            {
                inputForm.Width = 400;
                inputForm.Height = 300;
                inputForm.Text = "Enter Part Numbers (one per line)";

                TextBox textBox = new TextBox()
                {
                    Multiline = true,
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Vertical
                };
                inputForm.Controls.Add(textBox);

                Button okButton = new Button() { Text = "OK", Dock = DockStyle.Bottom };
                okButton.Click += (s, e) => inputForm.DialogResult = DialogResult.OK;
                inputForm.Controls.Add(okButton);

                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    return textBox.Text;
                }
            }
            return null;
        }
        private string PromptForOutputZipLocation()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Title = "Select Output ZIP Location";
                saveFileDialog.Filter = "ZIP files (*.zip)|*.zip";
                saveFileDialog.DefaultExt = "zip";
                saveFileDialog.AddExtension = true;

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return null;
        }
        private List<ProdFile> ScanFiles(string rootDirectory)
        {
            List<ProdFile> files = new List<ProdFile>();

            foreach (string filePath in Directory.GetFiles(rootDirectory, "*", SearchOption.AllDirectories))
            {
                try
                {
                    files.Add(new ProdFile(filePath)); // Updated to use ProdFile
                }
                catch
                {
                    // Skip invalid files
                }
            }
            return files;
        }




    }
}
