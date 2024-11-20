using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPDM.Interop.epdm;

namespace XpandNPIManager
{
    public partial class Form1 : Form
    {
        private Stopwatch stopwatch;

        public Form1()
        {
            InitializeComponent();
            stopwatch = new Stopwatch();
            progressBar.Visible = false;
            lblStatus.Visible = false;
        }

        private async void buttonListFiles_Click(object sender, EventArgs e)
        {
            try
            {
                string vaultName = PromptForVaultName();
                if (string.IsNullOrEmpty(vaultName))
                {
                    MessageBox.Show("Vault name cannot be empty. Operation cancelled.");
                    return;
                }

                IEdmVault5 vault = new EdmVault5();
                vault.LoginAuto(vaultName, this.Handle.ToInt32());

                if (!vault.IsLoggedIn)
                {
                    MessageBox.Show("Failed to log in to the vault.");
                    return;
                }

                IEdmFolder5 rootFolder = vault.RootFolder;
                string selectedFolderPath = BrowseForFolder(rootFolder.LocalPath);
                if (string.IsNullOrEmpty(selectedFolderPath))
                {
                    MessageBox.Show("No folder selected. Operation cancelled.");
                    return;
                }

                IEdmFolder5 selectedFolder = vault.GetFolderFromPath(selectedFolderPath);
                if (selectedFolder == null)
                {
                    MessageBox.Show("Selected folder is not part of the vault.");
                    return;
                }

                progressBar.Visible = true;
                lblStatus.Visible = true;
                progressBar.Value = 0;
                lblStatus.Text = "Scanning files...";
                stopwatch.Start();

                string listsFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Lists");
                Directory.CreateDirectory(listsFolderPath);

                string timestamp = DateTime.Now.ToString("dd-MM-yyyy-HH-mm");
                string csvFileName = $"Files{timestamp}.csv";
                string csvFilePath = Path.Combine(listsFolderPath, csvFileName);

                int totalFiles = await Task.Run(() => CountTotalFiles(selectedFolder));
                progressBar.Maximum = totalFiles;

                await Task.Run(() => ListFilesRecursively(selectedFolder, csvFilePath, totalFiles));

                stopwatch.Stop();
                progressBar.Visible = false;
                lblStatus.Visible = false;

                MessageBox.Show($"File list has been saved to: {csvFilePath}");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + "\n" + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private int CountTotalFiles(IEdmFolder5 folder)
        {
            int count = 0;
            IEdmPos5 filePos = folder.GetFirstFilePosition();
            while (!filePos.IsNull)
            {
                folder.GetNextFile(filePos);
                count++;
            }

            IEdmPos5 subFolderPos = folder.GetFirstSubFolderPosition();
            while (!subFolderPos.IsNull)
            {
                IEdmFolder5 subFolder = folder.GetNextSubFolder(subFolderPos);
                count += CountTotalFiles(subFolder);
            }

            return count;
        }

        private void ListFilesRecursively(IEdmFolder5 folder, string csvFilePath, int totalFiles)
        {
            int processedFiles = 0;

            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                writer.WriteLine("File Name,File Path,Last Modified,IsProd,Revision,Part Number");
                ListFiles(folder, writer, ref processedFiles, totalFiles);
            }
        }

        private void ListFiles(IEdmFolder5 folder, StreamWriter writer, ref int processedFiles, int totalFiles)
        {
            IEdmPos5 filePos = folder.GetFirstFilePosition();
            while (!filePos.IsNull)
            {
                IEdmFile5 file = folder.GetNextFile(filePos);
                string fileName = file.Name;
                Console.WriteLine($"Processing file: {fileName}"); // Debug output

                // Filter only PDF and Parasolid (.x_t) files
                if (fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) || fileName.EndsWith(".x_t", StringComparison.OrdinalIgnoreCase))
                {
                    string filePath = file.GetLocalPath(folder.ID);
                    DateTime lastModified = DateTime.MinValue;

                    // Check if the file is cached locally
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            lastModified = File.GetLastWriteTime(filePath);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error getting last modified date for {filePath}: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"File not cached locally: {filePath}");
                    }

                    // Initialize default values
                    string isProd = "0";
                    string revision = "";
                    string partNumber = fileName;

                    // Split the filename by underscores and dots
                    string[] parts = fileName.Split(new[] { '_', '.' }, StringSplitOptions.RemoveEmptyEntries);

                    // Loop through parts to find the revision and construct the part number
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (parts[i].StartsWith("Rev") && parts[i].Length == 4 && char.IsUpper(parts[i][3]))
                        {
                            // Found the revision
                            revision = parts[i].Substring(3);
                            isProd = "1";

                            // Construct the part number by excluding the revision and file extension
                            partNumber = string.Join("_", parts, 0, i);
                            break;
                        }
                    }

                    Console.WriteLine($"Extracted part number: {partNumber}, revision: {revision}, IsProd: {isProd}"); // Debug output

                    // Write the file details to the CSV
                    writer.WriteLine($"{fileName},{filePath},{lastModified:yyyy-MM-dd HH:mm:ss},{isProd},{revision},{partNumber}");
                }

                processedFiles++;
                UpdateProgress(processedFiles, totalFiles);
            }

            IEdmPos5 subFolderPos = folder.GetFirstSubFolderPosition();
            while (!subFolderPos.IsNull)
            {
                IEdmFolder5 subFolder = folder.GetNextSubFolder(subFolderPos);
                ListFiles(subFolder, writer, ref processedFiles, totalFiles);
            }
        }


        private void UpdateProgress(int processedFiles, int totalFiles)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateProgress(processedFiles, totalFiles)));
                return;
            }

            progressBar.Value = processedFiles;

            TimeSpan elapsed = stopwatch.Elapsed;
            double estimatedTotalSeconds = (elapsed.TotalSeconds / processedFiles) * totalFiles;
            TimeSpan estimatedTotalTime = TimeSpan.FromSeconds(estimatedTotalSeconds);
            TimeSpan remainingTime = estimatedTotalTime - elapsed;

            string elapsedTimeFormatted = elapsed.ToString(@"hh\:mm\:ss");
            string remainingTimeFormatted = remainingTime.ToString(@"hh\:mm\:ss");

            lblStatus.Text = $"Processed {processedFiles}/{totalFiles}. " +
                             $"Elapsed time: {elapsedTimeFormatted}. " +
                             $"Estimated remaining time: {remainingTimeFormatted}.";
        }

        private string BrowseForFolder(string rootFolderPath)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder inside the vault:";
                folderDialog.SelectedPath = rootFolderPath;
                folderDialog.ShowNewFolderButton = false;

                return folderDialog.ShowDialog() == DialogResult.OK ? folderDialog.SelectedPath : null;
            }
        }

        private string PromptForVaultName()
        {
            return Microsoft.VisualBasic.Interaction.InputBox("Enter the PDM Vault name:", "Vault Name", "");
        }
    }
}
