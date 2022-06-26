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
                this.outputTextBox.Text = string.Empty;

                this.selectedPath = fbd.SelectedPath;
                this.files = Directory.GetFiles(this.selectedPath);
                print("Files found: " + this.files.Length.ToString());
                print("Directory changed to: " + this.selectedPath);

                // Print each file in the directory. TODO: May not be needed later.
                int count = 1;
                foreach(string file in this.files){
                    string[] fileNameDir = file.Split('\\');
                    print(count + ". " + fileNameDir[fileNameDir.Length-1]);
                    count++;
                }
            }
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
            this.outputTextBox.Text = string.Empty;
        }
    }
}
