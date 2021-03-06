using Aspose.Cells;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvoiceToolsForm
{
    public partial class MainWindow : Form
    {
        // Path for the last folder chosen.
        private string selectedPath;

        // List of filepaths containing all files after a directory has been chosen, or if files have been renamed.
        private string[] allFilesInDirectory;

        // List of all OLD file paths.
        private List<string> oldFilePaths = new List<string>();

        // List of renamed file paths.
        private List<string> newFilePaths = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        // Choose a working directory
        private void DirectoryButton_Click(object sender, EventArgs e)
        {
            this.oldFilePaths.Clear();
            this.newFilePaths.Clear();

            this.selectedPath = string.Empty;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            // Report the number of files found in directory.
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                // reset print window
                clearOutput();

                this.selectedPath = fbd.SelectedPath;
                this.allFilesInDirectory = Directory.GetFiles(this.selectedPath);
                print("Directory changed to: " + this.selectedPath);

                foreach(string filePath in this.allFilesInDirectory){

                    int indexStart = filePath.Length - ".xlsx".Length;

                    if (filePath.Substring(indexStart, ".xlsx".Length).Equals(".xlsx"))
                    {
                        print(filePath);
                        this.oldFilePaths.Add(filePath);
                        this.newFilePaths.Add(filePath);
                    }
                }

                print(this.oldFilePaths.Count + " Excel files found.");
                return;
            }
            print("Error reading directory.");
        }

        // Process the xls files to find the invoice number, building address and first name of person that it should send to.
        // Afterwards, close file if needed and rename it given the found attributes.
        private void RenameButton_Click(object sender, EventArgs e)
        {

            // Check to see if there is an excel file open.
            // Prompt user to close all instances before proceeding.
            foreach(string filePath in this.oldFilePaths)
            {
                if (filePath.Contains("~$"))
                {
                    print("Close ALL EXCEL FILES");
                    return;
                }
            }

            // For each excel file, we want to rename it.
            int renameCount = 0;
            foreach(string oldPath in this.oldFilePaths)
            {
                // Load the excel file worksheet. (invoice)
                FileInfo fileInfo = new FileInfo(oldPath);
                ExcelPackage pack = new ExcelPackage(fileInfo);
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = pack.Workbook.Worksheets.FirstOrDefault();

                // Grab the relevant information from the worksheet.
                string to = worksheet.Cells[8, 1].Value.ToString();
                string location = worksheet.Cells[9, 1].Value.ToString();
                string[] invoiceNumSplit = worksheet.Cells[7, 8].Value.ToString().Split(' ');
                string invoiceNum = invoiceNumSplit[invoiceNumSplit.Length - 1];

                // Construct the new invoice name: ' Invoice #### - FirstName LastName - Building.xlsx '
                string newName = "Invoice " + invoiceNum + " - " + to + ".xlsx";

                // Copy the first part of the old path and add the new name to it.
                string newPath = string.Empty;
                string[] oldPathSplit = oldPath.Split('\\');

                for (int i = 0; i < oldPathSplit.Length - 1; i++)
                {
                    newPath += oldPathSplit[i] + '\\';
                }

                newPath += newName;

                newFilePaths.Add(newPath);

                // Rename the old file path with the new filepath.
                File.Move(oldPath, newPath);
                renameCount++;
            }

            // clean up after the files have been renamed
            this.oldFilePaths.Clear();
            print("Files renamed: " + renameCount);
            Console.WriteLine("Number of files in the member variable: " + this.newFilePaths.Count);
        }

        private void UpdateRecordButton_Click(object sender, EventArgs e)
        {

        }

        private void PdfPrintButton_Click(object sender, EventArgs e)
        {
            foreach(string path in newFilePaths)
            {
                Workbook workbook = new Workbook(path);

                string pdfFilename = path.TrimEnd(new char[] { 'x', 'l', 's', 'x' }) + "pdf";

                workbook.Save(pdfFilename, SaveFormat.Pdf);
            }
        }

        private void EmailButton_Click(object sender, EventArgs e)
        {

        }

        private void print(string s)
        {
            string prev = this.outputTextBox.Text;
            prev += s + "\n";
            this.outputTextBox.Text = prev;
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            clearOutput();
        }

        private void clearOutput()
        {
            this.outputTextBox.Text = string.Empty;
        }
    }
}
