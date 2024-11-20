using System;
using System.Diagnostics;
using System.IO;
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
            await ExecuteListAllFiles();
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
                    DateTime lastModified = DateTime.MinValue;

                    // Declare a variable to hold the handle
                    int parentWindowHandle = 0;

                    // Safely retrieve the handle from the UI thread
                    if (InvokeRequired)
                    {
                        Invoke(new Action(() => parentWindowHandle = this.Handle.ToInt32()));
                    }
                    else
                    {
                        parentWindowHandle = this.Handle.ToInt32();
                    }

                    try
                    {
                        Console.WriteLine($"Attempting to download file: {filePath}");
                        Console.WriteLine($"File Name: {file.Name}, Folder ID: {folder.ID}");

                        // Call GetFileCopy with adjusted parameters
                        object versionNoOrRevisionName = 0; // Get the latest version
                        object pathOrFolderID = folder.ID; // Use the folder ID to specify the download location
                        file.GetFileCopy(parentWindowHandle, ref versionNoOrRevisionName, ref pathOrFolderID, (int)EdmGetFlag.EdmGet_MakeReadOnly, "");

                        if (File.Exists(filePath))
                        {
                            Console.WriteLine($"File successfully downloaded: {filePath}");
                            lastModified = File.GetLastWriteTime(filePath);
                            Console.WriteLine($"Date Modified: {lastModified}");
                        }
                        else
                        {
                            Console.WriteLine($"File not found in local cache after download attempt: {filePath}");
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        Console.WriteLine($"Error during file download: HRESULT = 0x{ex.ErrorCode:X}, Message: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"General error during file download: {ex.Message}");
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
            double estimatedTotalSeconds = (elapsed.TotalSeconds / processedFiles) * totalFiles;
            TimeSpan estimatedTotalTime = TimeSpan.FromSeconds(estimatedTotalSeconds);
            TimeSpan remainingTime = estimatedTotalTime - elapsed;

            lblStatus.Text = $"Processed {processedFiles}/{totalFiles}. Elapsed: {elapsed:hh\\:mm\\:ss}. Remaining: {remainingTime:hh\\:mm\\:ss}.";
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

        private async Task ExecuteListAllFiles()
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
                    Console.WriteLine("Failed to log in to the vault.");
                    MessageBox.Show("Failed to log in to the vault.");
                    return;
                }
                Console.WriteLine($"Successfully logged into vault: {vaultName}");

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
    }
}
