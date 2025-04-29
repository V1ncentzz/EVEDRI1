using AForge;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using TheArtOfDevHtmlRenderer.Adapters;

namespace EVEDRI1
{
    public partial class Dashboard : Form
    {
        Workbook book = new Workbook();
        string name = string.Empty;
        private Form currentForm = null;

        public Dashboard()
        {
            InitializeComponent();
            picUser.Visible = true;
            lblUserfulname.Visible = true;
            pnlsettings.Visible = false;
            //.Text = showCount("Count" + num);
            //Mylogs mylogs = new Mylogs();
            //mylogs.insertlog( lblUserfulname.Text , "message");

        }

        bool sidebarExpand = true;
        bool pnlDashboardExpand = true;
        private void timer1_Tick(object sender, EventArgs e)
        {
           if (sidebarExpand)
            {
                Sidebar.Width -= 10;
                if(Sidebar.Width <= 63)
                {
                    sidebarExpand = false;
                    Sidebartranstion.Stop();
                }
            }
            else
            {
                Sidebar.Width += 10;
                if (Sidebar.Width >= 282)
                {
                    sidebarExpand = true;
                    Sidebartranstion.Stop();
                }
            }
        }

        private void LoadpanelForm(Form panelForm)
        {
            if (currentForm != null)
            {
                currentForm.Close(); 
            }
            currentForm = panelForm; 
            panelForm.TopLevel = false;
            panelForm.FormBorderStyle = FormBorderStyle.None; 
            panelForm.Dock = DockStyle.Fill;
            pnllDashboard.Controls.Add(panelForm); 
            pnllDashboard.Tag = panelForm; 
            panelForm.BringToFront(); 
            panelForm.Show(); 
        }

        //public int showcount(int c, string val)
        //{
             
        //}


        private void btnOptions_Click(object sender, EventArgs e)
        {
            Sidebartranstion.Start();
            picUser.Visible = false;
            lblUserfulname.Visible = false;
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            pnlsettings.Visible = true;
        }

        private void btnCLose_Click(object sender, EventArgs e)
        {
            pnlsettings.Visible = false;
        }


        private void btnDashboard_Click(object sender, EventArgs e)
        {
            LoadpanelForm(new Dashboard());
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private Rectangle prevBounds;
        private bool isMaximized = false;
        private void btnMaximize_Click(object sender, EventArgs e)
        {
            if (!isMaximized)
            {
                prevBounds = this.Bounds;
                this.WindowState = FormWindowState.Maximized;
                isMaximized = true;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.Bounds = prevBounds;
                isMaximized = false;
            }
        }

        private void btnClosewindow_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnActivestudent_Click(object sender, EventArgs e)
        {
           LoadpanelForm(new Form2());
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
          LoadpanelForm(new Logs());
        }

        private void btnfromafile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtPicture.Text = openFileDialog.FileName;
            }
        }

        private void pnllDashboard_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            login.Show();
            this.Close();
        }
    }
}
