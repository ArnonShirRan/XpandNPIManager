using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPILib;

namespace XpandNPIManager
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                var files = new Files();

                // Define the folder path to scan and the CSV output path
                string folderPath = @"\\PDM\ERP-Data\Files"; // Replace with your test folder path
                string outputPath = @"\\PDM\ERP-Data\Files\FileList.csv"; // Replace with your desired CSV path

                // Scan the folder recursively and add files to the Files object
                var scannedFiles = FolderScanner.ScanFolderRecursively(folderPath);
                files.AddFiles(scannedFiles);

                // Display the scanned files
                FileDisplay.DisplayFiles(files.FileList);

                // Display skipped files
                ScanLogger.DisplaySkippedFiles();

                // Export the scanned files to CSV
                CSVCreator.CreateFilesCSV(outputPath, files);

                Console.WriteLine("Process completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
