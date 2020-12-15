using System.IO;
using Tools;

namespace OLECalApplication
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
        private FileSystemWatcher fileSystemWatcher1;

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
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "Form1";
        }

        private void InitializeComponent(string path, string username, string password, string domainName)
        {

           // using (new Impersonator(username, domainName, password))
            {
                this.fileSystemWatcher1 = new System.IO.FileSystemWatcher();
                ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).BeginInit();
                this.SuspendLayout();
                // 
                // fileSystemWatcher1
                // 
                this.fileSystemWatcher1.EnableRaisingEvents = true;
                this.fileSystemWatcher1.SynchronizingObject = this;
                this.fileSystemWatcher1.EnableRaisingEvents = true;
                //this.fileSystemWatcher1.Path = "\\\\SRKS_TECH-1\\Users\\Public\\FTP\\AutoUpdatedFile\\\\";
                this.fileSystemWatcher1.Path = path;
                //\\SRKS_TECH-2\Users\Tech-2\J\2016-03-16
                this.fileSystemWatcher1.SynchronizingObject = this;
                this.fileSystemWatcher1.Created += new System.IO.FileSystemEventHandler(this.fileSystemWatcher1_Created);
                // 
                // Form1
                // 
                this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
                this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                this.ClientSize = new System.Drawing.Size(284, 261);
                this.Name = "Form1";
                this.Text = "Form1";
                ((System.ComponentModel.ISupportInitialize)(this.fileSystemWatcher1)).EndInit();
                this.ResumeLayout(false);
                this.ShowInTaskbar = false;
                this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                this.ShowIcon = false;
            }

        }

        #endregion
    }
}

