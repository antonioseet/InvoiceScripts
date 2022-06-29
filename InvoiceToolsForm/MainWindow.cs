using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvoiceToolsForm
{
    public partial class MainWindow : Form
    {
        private string selectedPath;
        private string[] files;
        private List<string> xlsxFiles;

        public MainWindow()
        {
            InitializeComponent();
        }

        // Choose a working directory
        private void DirectoryButton_Click(object sender, EventArgs e)
        {
            this.selectedPath = string.Empty;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();

            // Report the number of files found in directory.
            if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                // reset print window
                clearOutput();

                this.selectedPath = fbd.SelectedPath;
                this.files = Directory.GetFiles(this.selectedPath);
                print("Directory changed to: " + this.selectedPath);

                this.xlsxFiles = new List<string>();
                foreach(string fileName in this.files){

                    int indexStart = fileName.Length - ".xlsx".Length;

                    if (fileName.Substring(indexStart, ".xlsx".Length).Equals(".xlsx"))
                    {
                        print(fileName);
                        this.xlsxFiles.Add(fileName);
                    }
                }

                print(this.xlsxFiles.Count + " Excel files found.");
                return;
            }
            print("Error reading directory.");
        }

        // Process the xls files to find the invoice number, building address and first name of person that it should send to.
        // Afterwards, close file if needed and rename it given the found attributes.
        private void RenameButton_Click(object sender, EventArgs e)
        {

        }

        private void UpdateRecordButton_Click(object sender, EventArgs e)
        {

        }

        private void PdfPrintButton_Click(object sender, EventArgs e)
        {

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
