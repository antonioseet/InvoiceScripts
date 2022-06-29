
namespace InvoiceToolsForm
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.directoryButton = new System.Windows.Forms.Button();
            this.renameButton = new System.Windows.Forms.Button();
            this.updateRecordButton = new System.Windows.Forms.Button();
            this.pdfPrintButton = new System.Windows.Forms.Button();
            this.emailButton = new System.Windows.Forms.Button();
            this.outputTextBox = new System.Windows.Forms.RichTextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ClearButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // directoryButton
            // 
            this.directoryButton.Location = new System.Drawing.Point(12, 12);
            this.directoryButton.Name = "directoryButton";
            this.directoryButton.Size = new System.Drawing.Size(204, 77);
            this.directoryButton.TabIndex = 0;
            this.directoryButton.Text = "Choose Directory";
            this.directoryButton.UseVisualStyleBackColor = true;
            this.directoryButton.Click += new System.EventHandler(this.DirectoryButton_Click);
            // 
            // renameButton
            // 
            this.renameButton.Location = new System.Drawing.Point(12, 95);
            this.renameButton.Name = "renameButton";
            this.renameButton.Size = new System.Drawing.Size(204, 77);
            this.renameButton.TabIndex = 1;
            this.renameButton.Text = "Rename All Invoices";
            this.renameButton.UseVisualStyleBackColor = true;
            this.renameButton.Click += new System.EventHandler(this.RenameButton_Click);
            // 
            // updateRecordButton
            // 
            this.updateRecordButton.Location = new System.Drawing.Point(12, 178);
            this.updateRecordButton.Name = "updateRecordButton";
            this.updateRecordButton.Size = new System.Drawing.Size(204, 77);
            this.updateRecordButton.TabIndex = 2;
            this.updateRecordButton.Text = "Update Records";
            this.updateRecordButton.UseVisualStyleBackColor = true;
            this.updateRecordButton.Click += new System.EventHandler(this.UpdateRecordButton_Click);
            // 
            // pdfPrintButton
            // 
            this.pdfPrintButton.Location = new System.Drawing.Point(12, 261);
            this.pdfPrintButton.Name = "pdfPrintButton";
            this.pdfPrintButton.Size = new System.Drawing.Size(204, 77);
            this.pdfPrintButton.TabIndex = 3;
            this.pdfPrintButton.Text = "Print To PDF";
            this.pdfPrintButton.UseVisualStyleBackColor = true;
            this.pdfPrintButton.Click += new System.EventHandler(this.PdfPrintButton_Click);
            // 
            // emailButton
            // 
            this.emailButton.Location = new System.Drawing.Point(12, 344);
            this.emailButton.Name = "emailButton";
            this.emailButton.Size = new System.Drawing.Size(204, 77);
            this.emailButton.TabIndex = 4;
            this.emailButton.Text = "Email to addresses";
            this.emailButton.UseVisualStyleBackColor = true;
            this.emailButton.Click += new System.EventHandler(this.EmailButton_Click);
            // 
            // outputTextBox
            // 
            this.outputTextBox.Location = new System.Drawing.Point(6, 19);
            this.outputTextBox.Name = "outputTextBox";
            this.outputTextBox.Size = new System.Drawing.Size(528, 401);
            this.outputTextBox.TabIndex = 5;
            this.outputTextBox.Text = "";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.outputTextBox);
            this.groupBox1.Location = new System.Drawing.Point(248, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(540, 426);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Output:";
            // 
            // ClearButton
            // 
            this.ClearButton.Location = new System.Drawing.Point(713, 445);
            this.ClearButton.Name = "ClearButton";
            this.ClearButton.Size = new System.Drawing.Size(75, 23);
            this.ClearButton.TabIndex = 7;
            this.ClearButton.Text = "clear";
            this.ClearButton.UseVisualStyleBackColor = true;
            this.ClearButton.Click += new System.EventHandler(this.ClearButton_Click);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 508);
            this.Controls.Add(this.ClearButton);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.emailButton);
            this.Controls.Add(this.pdfPrintButton);
            this.Controls.Add(this.updateRecordButton);
            this.Controls.Add(this.renameButton);
            this.Controls.Add(this.directoryButton);
            this.Name = "MainWindow";
            this.Text = "Invoice Tools";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button directoryButton;
        private System.Windows.Forms.Button renameButton;
        private System.Windows.Forms.Button updateRecordButton;
        private System.Windows.Forms.Button pdfPrintButton;
        private System.Windows.Forms.Button emailButton;
        private System.Windows.Forms.RichTextBox outputTextBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button ClearButton;
    }
}

