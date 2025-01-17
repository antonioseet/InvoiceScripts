﻿using OfficeOpenXml;
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

/* List of Ideas:
 * - Create a folder to store a sample invoice
 * - Use the sample invoice to create a new invoice in the folder associated with the current date
 *      for example, if the current date is 12/3/2023, create a folder '2023-12' and copy the sample invoice into it.
 *      We can create a button for this function for now.
 *      
 * - If we try to rename, and the file is a sample file (no changes), then we should not rename it. and we should not print it.
 * - try to skip over files that have already been printed to PDF. (this may not work because maybe we change the excel sheet after printing it)
 *      which means that if we skip over it, we will not print the post-print modified files.
 *      
 * - When the app starts, instead of having the user choose a directory, auto detect the directory we want to use.
 *      Usually this is done by checking the date.
 *      
 * - Auto populate Invoice Number based on the last invoice number used.
 */


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

            this.fbd.SelectedPath = "C:\\Users\\aabar\\OneDrive\\GC Documents\\Invoices";

            DialogResult result = fbd.ShowDialog();

            // Report the number of files found in directory.
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(this.fbd.SelectedPath))
            {
                updateWindowWithFilenames(true /* updateNewFilePaths */);
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
            clearOutput();

            selectedPath = fbd.SelectedPath;

            // get all excel files in the directory
            allFilesInDirectory = Directory.GetFiles(selectedPath, "*.xlsx");

            foreach (string filePath in allFilesInDirectory)
            {
                oldFilePaths.Add(filePath);

                if (updateNewFilePaths)
                    newFilePaths.Add(filePath);
            }

            print(oldFilePaths.Count + " Excel files found.");
            printStatus();
        }

        // Checks if files are open by checking if any of the filepaths contain '~$'
        // This prevents the program from trying to rename files that are open.
        // TODO: this stopped working for some reason. look into it.
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

                // Grab the name and location from the worksheet.
                // Cell[8,1] = A8
                // Cell[9,1] = A9
                string toName = worksheet.Cells[8, 1].Value.ToString();
                string location = worksheet.Cells[9, 1].Value.ToString();

                // TODO: Handle 'Quotes/Estimates'
                // Estimates will not have invoice number associated with it, causing a break.
                string newName;
                string invoiceNumber = "0";
                string value = worksheet.Cells[1, 7].Value.ToString();
                bool isInvoice = value == "INVOICE";

                if (isInvoice)
                {
                    string[] invoiceNumSplit = worksheet.Cells[7, 8].Value.ToString().Split(' ');
                    invoiceNumber = invoiceNumSplit[invoiceNumSplit.Length - 1];
                    
                    // Construct the new invoice name: ' Invoice #### - FirstName LastName - Building.xlsx '
                    newName = getRename(invoiceNumber, toName, location);
                }
                else
                {
                    newName = getRename(toName, location);
                }


                

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
                // TODO: Cannot create a file when that file already exists. Need to delete the old file first. OR RENAME THE OLD FILEs FIRST.
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
            // TODO: Refresh list of files in directory because they may have changed since the last time we checked.
            // This causes a crash because we try to print pdf files for excel files that don't exist anymore.
            // repro: rename files, then delete one, then try to print pdfs.
            // TODO: Whne continuously pressing this button, the pdf count does not get reset. will continually go up, even if the files are deleted.

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
        private string getRename(string name, string location)
        {
            if (renameCheckbox.Checked)
                return "Quote - " + name + " - " + location + ".xlsx"; ;

            return "Quote - " + location + ".xlsx"; ;
        }

    }
}
