using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
            await ExecuteListAllFiles();
        }

        private async void btnGetFileLocations_Click(object sender, EventArgs e)
        {
            // Ensure ExecuteListAllFiles completes successfully
            bool listCreationSucceeded = await ExecuteListAllFiles();
            if (listCreationSucceeded)
            {
                await ProcessFileLocationsAsync();
            }
        }

        private async Task<bool> ExecuteListAllFiles()
        {
            try
            {
                string vaultName = PromptForVaultName();
                if (string.IsNullOrEmpty(vaultName))
                {
                    MessageBox.Show("Vault name cannot be empty. Operation cancelled.");
                    return false;
                }

                IEdmVault5 vault = new EdmVault5();
                vault.LoginAuto(vaultName, this.Handle.ToInt32());

                if (!vault.IsLoggedIn)
                {
                    MessageBox.Show("Failed to log in to the vault.");
                    return false;
                }

                IEdmFolder5 rootFolder = vault.RootFolder;
                string selectedFolderPath = BrowseForFolder(rootFolder.LocalPath);
                if (string.IsNullOrEmpty(selectedFolderPath))
                {
                    MessageBox.Show("No folder selected. Operation cancelled.");
                    return false;
                }

                IEdmFolder5 selectedFolder = vault.GetFolderFromPath(selectedFolderPath);
                if (selectedFolder == null)
                {
                    MessageBox.Show("Selected folder is not part of the vault.");
                    return false;
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
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
                return false;
            }
        }

        private async Task ProcessFileLocationsAsync()
        {
            // Prompt the user for part numbers
            string inputPartNumbers = PromptForPartNumbers();
            if (string.IsNullOrWhiteSpace(inputPartNumbers))
            {
                MessageBox.Show("No part numbers entered. Operation cancelled.");
                return;
            }

            // Split the input into individual part numbers
            var partNumbers = inputPartNumbers.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Find the most recent file in the Lists folder
            string latestCsvFile = GetMostRecentCsv("Lists");
            if (string.IsNullOrEmpty(latestCsvFile))
            {
                MessageBox.Show("No CSV file found in the Lists folder. Operation cancelled.");
                return;
            }

            // Process the most recent file to find the most updated files for the part numbers
            var pnRecords = RetrieveMostUpdatedFilesFromCsv(latestCsvFile, partNumbers);

            // Save the results to a new CSV in the Files folder
            SaveToPnFilesCsv(pnRecords);

            MessageBox.Show("File locations have been processed and saved.");
        }

        private string PromptForPartNumbers()
        {
            using (var form = new Form())
            {
                form.Text = "Enter Part Numbers";
                form.Size = new System.Drawing.Size(400, 300);

                var textBox = new TextBox
                {
                    Multiline = true,
                    Dock = DockStyle.Fill,
                    ScrollBars = ScrollBars.Vertical
                };

                var okButton = new Button
                {
                    Text = "OK",
                    Dock = DockStyle.Bottom
                };

                okButton.Click += (s, e) => form.DialogResult = DialogResult.OK;

                form.Controls.Add(textBox);
                form.Controls.Add(okButton);

                return form.ShowDialog() == DialogResult.OK ? textBox.Text : null;
            }
        }

        private string GetMostRecentCsv(string folderName)
        {
            string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, folderName);
            if (!Directory.Exists(folderPath))
            {
                return null;
            }

            return Directory.GetFiles(folderPath, "*.csv")
                .OrderByDescending(File.GetLastWriteTime)
                .FirstOrDefault();
        }

        private List<PnRecord> RetrieveMostUpdatedFilesFromCsv(string csvPath, string[] partNumbers)
        {
            var records = new List<PnRecord>();
            var lines = File.ReadAllLines(csvPath);

            foreach (var partNumber in partNumbers)
            {
                string mostUpdatedXtFile = null;
                string mostUpdatedPdfFile = null;
                DateTime mostRecentDate = DateTime.MinValue;

                foreach (var line in lines.Skip(1)) // Skip header
                {
                    var columns = line.Split(',');
                    if (columns.Length < 6) continue;

                    string fileName = columns[0];
                    string filePath = columns[1];
                    DateTime fileDate = DateTime.Parse(columns[2]);
                    string currentPartNumber = columns[5];

                    if (!partNumber.Equals(currentPartNumber, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (filePath.EndsWith(".x_t", StringComparison.OrdinalIgnoreCase) && fileDate > mostRecentDate)
                    {
                        mostUpdatedXtFile = filePath;
                        mostRecentDate = fileDate;
                    }

                    if (filePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) && fileDate > mostRecentDate)
                    {
                        mostUpdatedPdfFile = filePath;
                        mostRecentDate = fileDate;
                    }
                }

                records.Add(new PnRecord
                {
                    PartNumber = partNumber,
                    MostUpdatedXtFile = mostUpdatedXtFile,
                    MostUpdatedPdfFile = mostUpdatedPdfFile
                });
            }

            return records;
        }

        private void SaveToPnFilesCsv(List<PnRecord> pnRecords)
        {
            string filesFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files");
            Directory.CreateDirectory(filesFolderPath);

            string outputCsvPath = Path.Combine(filesFolderPath, $"PNFiles_{DateTime.Now:dd-MM-yyyy-HH-mm}.csv");

            using (var writer = new StreamWriter(outputCsvPath))
            {
                writer.WriteLine("Part Number,Most Updated x_t File,Most Updated pdf File");
                foreach (var record in pnRecords)
                {
                    writer.WriteLine($"{record.PartNumber},{record.MostUpdatedXtFile},{record.MostUpdatedPdfFile}");
                }
            }
        }

        public class PnRecord
        {
            public string PartNumber { get; set; }
            public string MostUpdatedXtFile { get; set; }
            public string MostUpdatedPdfFile { get; set; }
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
                Console.WriteLine($"Processing file: {fileName}");

                if (fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) || fileName.EndsWith(".x_t", StringComparison.OrdinalIgnoreCase))
                {
                    string filePath = file.GetLocalPath(folder.ID);
                    DateTime lastModified = File.GetLastWriteTime(filePath);

                    string isProd = "0";
                    string revision = "";
                    string partNumber = fileName;

                    string[] parts = fileName.Split(new[] { '_', '.' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (parts[i].StartsWith("Rev") && parts[i].Length == 4 && char.IsUpper(parts[i][3]))
                        {
                            revision = parts[i].Substring(3);
                            isProd = "1";
                            partNumber = string.Join("_", parts, 0, i);
                            break;
                        }
                    }

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

            lblStatus.Text = $"Processed {processedFiles}/{totalFiles}. Elapsed: {elapsed:hh\\:mm\\:ss}. Remaining: {remainingTime:hh\\:mm\\:ss}.";
        }

        private string BrowseForFolder(string rootFolderPath)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder inside the vault:";
                folderDialog.SelectedPath = rootFolderPath; // Set the initial directory to the root folder
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }

                return null;
            }
        }




        private string PromptForVaultName()
        {
            return Microsoft.VisualBasic.Interaction.InputBox("Enter the PDM Vault name:", "Vault Name", "");
        }
    }
}
