using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
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

        private void btnGetFileLocations_Click(object sender, EventArgs e)
        {
            // Step 1: ExecuteListAllFiles
            bool listCreationSucceeded = ExecuteListAllFiles();
            if (!listCreationSucceeded)
            {
                MessageBox.Show("Failed to create the file list. Operation cancelled.");
                return;
            }

            // Step 2: ProcessFileLocations
            bool fileLocationsProcessed = ProcessFileLocations();
            if (!fileLocationsProcessed)
            {
                MessageBox.Show("Failed to process file locations. Operation cancelled.");
                return;
            }

            // Step 3: Create ZIP from the latest file
            bool zipCreationSucceeded = CreateZipFromLatestFile();
            if (zipCreationSucceeded)
            {
                MessageBox.Show("ZIP file created successfully.");
            }
        }




        private bool ExecuteListAllFiles()
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

                string listsFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Lists");
                Directory.CreateDirectory(listsFolderPath);

                string timestamp = DateTime.Now.ToString("dd-MM-yyyy-HH-mm");
                string csvFileName = $"Files{timestamp}.csv";
                string csvFilePath = Path.Combine(listsFolderPath, csvFileName);

                int totalFiles = CountTotalFiles(selectedFolder);

                // Initialize progress tracking
                stopwatch.Restart();
                progressBar.Visible = true;
                lblStatus.Visible = true;
                progressBar.Maximum = totalFiles;
                progressBar.Value = 0;
                lblStatus.Text = "Starting scan...";

                ListFilesRecursively(selectedFolder, csvFilePath, totalFiles);

                stopwatch.Stop();
                progressBar.Visible = false;
                lblStatus.Text = "Scan completed.";

                MessageBox.Show($"File list has been saved to: {csvFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
                progressBar.Visible = false;
                lblStatus.Text = "Error occurred.";
                return false;
            }
        }







        private string PromptForVaultName()
        {
            return Microsoft.VisualBasic.Interaction.InputBox("Enter the PDM Vault name:", "Vault Name", "");
        }

        private bool ProcessFileLocations()
        {
            try
            {
                string inputPartNumbers = PromptForPartNumbers();
                if (string.IsNullOrWhiteSpace(inputPartNumbers))
                {
                    MessageBox.Show("No part numbers entered. Operation cancelled.");
                    return false;
                }

                var partNumbers = inputPartNumbers.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                string latestCsvFile = GetMostRecentCsv("Lists");
                if (string.IsNullOrEmpty(latestCsvFile))
                {
                    MessageBox.Show("No CSV file found in the Lists folder. Operation cancelled.");
                    return false;
                }

                var pnRecords = RetrieveMostUpdatedFilesFromCsv(latestCsvFile, partNumbers);
                SaveToPnFilesCsv(pnRecords);

                MessageBox.Show("File locations have been processed and saved.");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
                return false;
            }
        }


        private bool CreateZipFromLatestFile()
        {
            try
            {
                // Define the Files folder path
                string filesFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files");

                // Find the latest CSV file in the Files folder
                string latestCsvPath = Directory.GetFiles(filesFolderPath, "PNFiles_*.csv")
                                                .OrderByDescending(File.GetCreationTime)
                                                .FirstOrDefault();

                if (string.IsNullOrEmpty(latestCsvPath))
                {
                    MessageBox.Show("No CSV file found in the Files folder. Operation cancelled.");
                    return false;
                }

                // Read file paths from the CSV file
                List<string> filePaths = new List<string>();
                using (var reader = new StreamReader(latestCsvPath))
                {
                    // Skip the header row
                    reader.ReadLine();

                    // Read file paths from each record
                    while (!reader.EndOfStream)
                    {
                        string[] fields = reader.ReadLine()?.Split(',');
                        if (fields != null && fields.Length >= 3)
                        {
                            if (!string.IsNullOrEmpty(fields[1])) filePaths.Add(fields[1].Trim());
                            if (!string.IsNullOrEmpty(fields[2])) filePaths.Add(fields[2].Trim());
                        }
                    }
                }

                // Create the ZIP folder path
                string zipFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ZIP");
                Directory.CreateDirectory(zipFolderPath);

                // Define the ZIP file path
                string zipFileName = $"ProductionFiles_{DateTime.Now:yyyyMMdd_HHmmss}.zip";
                string zipFilePath = Path.Combine(zipFolderPath, zipFileName);

                // Create the ZIP file
                using (FileStream zipToOpen = new FileStream(zipFilePath, FileMode.Create))
                {
                    using (System.IO.Compression.ZipArchive archive = new System.IO.Compression.ZipArchive(zipToOpen, System.IO.Compression.ZipArchiveMode.Create))
                    {
                        foreach (string filePath in filePaths)
                        {
                            if (File.Exists(filePath))
                            {
                                string entryName = Path.GetFileName(filePath);
                                archive.CreateEntryFromFile(filePath, entryName);
                            }
                        }
                    }
                }

                MessageBox.Show($"ZIP file created successfully at: {zipFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while creating the ZIP file: {ex.Message}");
                return false;
            }
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

            lblStatus.Text = $"Completed scanning {processedFiles}/{totalFiles} files.";
            Application.DoEvents(); // Ensure UI updates after completion
        }




        private void buttonListFiles_Click(object sender, EventArgs e)
        {
            // Example logic for listing all files
            bool result = ExecuteListAllFiles();
            if (result)
            {
                MessageBox.Show("File list created successfully.");
            }
            else
            {
                MessageBox.Show("Failed to create file list.");
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
                    DateTime lastModified = DateTime.MinValue;

                    int parentWindowHandle = InvokeRequired
                        ? (int)Invoke(new Func<int>(() => this.Handle.ToInt32()))
                        : this.Handle.ToInt32();

                    try
                    {
                        Console.WriteLine($"Attempting to download file: {filePath}");
                        object versionNoOrRevisionName = 0; // Latest version
                        object pathOrFolderID = folder.ID;
                        file.GetFileCopy(parentWindowHandle, ref versionNoOrRevisionName, ref pathOrFolderID, (int)EdmGetFlag.EdmGet_MakeReadOnly, "");

                        if (File.Exists(filePath))
                        {
                            Console.WriteLine($"File successfully downloaded: {filePath}");
                            lastModified = File.GetLastWriteTime(filePath);
                        }
                        else
                        {
                            Console.WriteLine($"File not found after download attempt: {filePath}");
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        Console.WriteLine($"Error downloading file: HRESULT = 0x{ex.ErrorCode:X}, Message: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"General error downloading file: {ex.Message}");
                    }

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
            double estimatedTotalSeconds = (processedFiles > 0)
                ? (elapsed.TotalSeconds / processedFiles) * totalFiles
                : 0;
            TimeSpan estimatedTotalTime = TimeSpan.FromSeconds(estimatedTotalSeconds);
            TimeSpan remainingTime = estimatedTotalTime - elapsed;

            lblStatus.Text = $"Processed {processedFiles}/{totalFiles} files. " +
                             $"Elapsed: {elapsed:hh\\:mm\\:ss}. " +
                             $"Remaining: {remainingTime:hh\\:mm\\:ss}.";

            Application.DoEvents(); // Ensure UI updates during processing
        }







        private string BrowseForFolder(string rootFolderPath)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder inside the vault:";
                folderDialog.SelectedPath = rootFolderPath; // Set the initial directory to the vault root
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    return folderDialog.SelectedPath;
                }

                return null;
            }
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

        public class PnRecord
        {
            public string PartNumber { get; set; }
            public string MostUpdatedXtFile { get; set; }
            public string MostUpdatedPdfFile { get; set; }
        }
    }
}
