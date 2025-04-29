namespace EVEDRI1
{
    partial class Logs
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
            this.logsDatagidview = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.logsDatagidview)).BeginInit();
            this.SuspendLayout();
            // 
            // logsDatagidview
            // 
            this.logsDatagidview.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.logsDatagidview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logsDatagidview.Location = new System.Drawing.Point(0, 0);
            this.logsDatagidview.Name = "logsDatagidview";
            this.logsDatagidview.Size = new System.Drawing.Size(1414, 768);
            this.logsDatagidview.TabIndex = 0;
            this.logsDatagidview.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.logsDatagidview_CellContentClick);
            // 
            // Logs
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1414, 768);
            this.Controls.Add(this.logsDatagidview);
            this.Name = "Logs";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Logs";
            ((System.ComponentModel.ISupportInitialize)(this.logsDatagidview)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView logsDatagidview;
    }
}