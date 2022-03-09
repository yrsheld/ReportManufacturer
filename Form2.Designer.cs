namespace WindowsFormsApp14
{
    partial class Form2
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
            this.docDocumentViewer1 = new Spire.DocViewer.Forms.DocDocumentViewer();
            this.SuspendLayout();
            // 
            // docDocumentViewer1
            // 
            this.docDocumentViewer1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(229)))), ((int)(((byte)(229)))), ((int)(((byte)(229)))));
            this.docDocumentViewer1.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.docDocumentViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.docDocumentViewer1.EnableHandTools = false;
            this.docDocumentViewer1.Location = new System.Drawing.Point(0, 0);
            this.docDocumentViewer1.Name = "docDocumentViewer1";
            this.docDocumentViewer1.Size = new System.Drawing.Size(800, 450);
            this.docDocumentViewer1.TabIndex = 0;
            this.docDocumentViewer1.Text = "docDocumentViewer1";
            this.docDocumentViewer1.ToPdfParameterList = null;
            this.docDocumentViewer1.ZoomMode = Spire.DocViewer.Forms.ZoomMode.Default;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.docDocumentViewer1);
            this.Name = "Form2";
            this.Text = "Form2";
            this.ResumeLayout(false);

        }

        #endregion

        public Spire.DocViewer.Forms.DocDocumentViewer docDocumentViewer1;
    }
}