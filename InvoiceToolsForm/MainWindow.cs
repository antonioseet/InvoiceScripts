using OfficeOpenXml;
using Spire.Pdf.Exporting.XPS.Schema;
using Spire.Xls;
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
using System.Windows.Media.TextFormatting;
using System.Xml.Linq;

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

        // number of files printed.
        private int printCount = 0;

        // Instance of folder dialog to save filepath
        FolderBrowserDialog fbd;

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
            this.fbd = new FolderBrowserDialog();

            this.fbd.SelectedPath = "C:\\Users\\aabar\\OneDrive\\GC Documents\\Invoices\\2021-02";

            DialogResult result = fbd.ShowDialog();

            // Report the number of files found in directory.
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(this.fbd.SelectedPath))
            {
                updateWindowWithFilenames(true);
                enableButtons();

                return;
            }
            print("Error reading directory.");
        }

        private void enableButtons()
        {
            renameButton.Enabled = true;
            //updateRecordButton.Enabled = true;
            pdfPrintButton.Enabled = true;
            //emailButton.Enabled = true;
        }

        public void updateWindowWithFilenames(bool updateNewFilePaths)
        {
            // reset print window
            clearOutput();

            this.selectedPath = this.fbd.SelectedPath;
            this.allFilesInDirectory = Directory.GetFiles(this.selectedPath);
            print("Directory set to: " + this.selectedPath);

            foreach (string filePath in this.allFilesInDirectory)
            {

                int indexStart = filePath.Length - ".xlsx".Length;

                if (filePath.Substring(indexStart, ".xlsx".Length).Equals(".xlsx"))
                {
                    this.oldFilePaths.Add(filePath);

                    if (updateNewFilePaths)
                        this.newFilePaths.Add(filePath);
                }
            }

            print(this.oldFilePaths.Count + " Excel files found.");
            printStatus();
        }

        public bool filesAreOpen()
        {
            bool openFileFound = false;
            foreach (string filePath in this.oldFilePaths)
            {
                if (filePath.Contains("~$"))
                {
                    openFileFound = true;
                    break;
                }
            }
            return openFileFound;
        }

        // Process the xls files to find the invoice number, building address and first name of person that it should send to.
        // Afterwards, close file if needed and rename it given the found attributes.
        private void RenameButton_Click(object sender, EventArgs e)
        {
            // If we are trying to rename, these old filenames need to be removed to avoid trying to print them later.
            this.newFilePaths.Clear();
            this.oldFilePaths.Clear();

            // Check to see if there is an excel file open.
            // Prompt user to close all instances before proceeding.
            updateWindowWithFilenames(false);

            if (filesAreOpen())
            {
                print("CLOSE ALL EXCEL FILES AND TRY AGAIN.");
                return;
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
                string toName = worksheet.Cells[8, 1].Value.ToString();
                string location = worksheet.Cells[9, 1].Value.ToString();

                string[] invoiceNumSplit = worksheet.Cells[7, 8].Value.ToString().Split(' ');
                string invoiceNum = invoiceNumSplit[invoiceNumSplit.Length - 1];

                // Construct the new invoice name: ' Invoice #### - FirstName LastName - Building.xlsx '
                string newName = getRename(invoiceNum, toName, location);

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
            printStatus();
        }

        private void UpdateRecordButton_Click(object sender, EventArgs e)
        {
        }

        private void PdfPrintButton_Click(object sender, EventArgs e)
        {
            // TODO: To make things a little easier, try to delete all existing pdf files. 

            foreach(string path in newFilePaths)
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(path);

                string pdfFilename = path.TrimEnd(new char[] { 'x', 'l', 's', 'x' }) + "pdf";

                workbook.SaveToFile(pdfFilename, Spire.Xls.FileFormat.PDF);
                printCount++;
            }
            printStatus();
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

        private void printStatus()
        {
            print("");
            print("New Filepaths | Total: " + this.newFilePaths.Count + " | Total Printed: " + this.printCount); ;
            foreach (string path in this.newFilePaths)
            {
                print(" " + path);
            }
            print("");
        }

        /// <summary>
        /// Based on the checkbox at the bottom left corner, renames a document based on the two templates.
        /// </summary>
        /// <param name="num">Invoice Number</param>
        /// <param name="name">Invoicee</param>
        /// <param name="location">Job Site</param>
        /// <returns></returns>
        private string getRename(string num, string name, string location)
        {
            if (renameCheckbox.Checked)
                return "" + num + " - " + name + " - " + location + ".xlsx"; ;

            return "" + num + " - " + location + ".xlsx"; ;
        }

    }
}
