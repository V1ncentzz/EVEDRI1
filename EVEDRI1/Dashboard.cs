using Spire.Xls;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace EVEDRI1
{
    public partial class Dashboard : Form
    {
        //Workbook book = new Workbook();
        private Workbook workbook;
        private Worksheet worksheet;
        private string excelFilePath = @"C:\Users\HF\Documents\Ff\Book1.xlsx";
        string name = string.Empty;
        private Form currentForm = null;

        int activeCount = 0;
        int inactiveCount = 0;
        int maleCount = 0;
        int femaleCount = 0;
        int volleyballCount = 0;
        int basketballCount = 0;
        int tennisCount = 0;
        int soccerCount = 0;
        int badmintonCount = 0;
        int baseballCount = 0;
        int bsitCount = 0;
        int bstmCount = 0;
        int beedCount = 0;
        int orangeCount = 0;
        int blueCount = 0;
        int yellowCount = 0;
        int redCount = 0;

        public Dashboard()
        {
            InitializeComponent();
            Loadexceldata();
            Count();
            picUser.Visible = true;
            lblUserfulname.Visible = true;
            pnlsettings.Visible = false;

        }

        private void Loadexceldata()
        {
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\EVEDRI1\Book1.xlsx");
                Worksheet worksheet = workbook.Worksheets[0];
            }
            catch(Exception ex)
            {
                MessageBox.Show($"Error loading Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void Count()
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\source\repos\EVEDRI1\Book1.xlsx");
            if (worksheet != null)
            {
                int lastRow = worksheet.LastRow;
                if (lastRow < 2) return;

                for (int i = 2; i <= lastRow; i++)
                {
                    string status = worksheet.GetText(i, 1)?.Trim();
                    if (string.Equals(status, "Active", StringComparison.OrdinalIgnoreCase)) activeCount++;
                    else if (string.Equals(status, "Inactive", StringComparison.OrdinalIgnoreCase)) inactiveCount++;

                    string gender = worksheet.GetText(i, 2)?.Trim();
                    if (string.Equals(gender, "Male", StringComparison.OrdinalIgnoreCase)) maleCount++;
                    else if (string.Equals(gender, "Female", StringComparison.OrdinalIgnoreCase)) femaleCount++;

                    // Count Hobbies (inside the loop)
                    if (string.Equals(worksheet.GetText(i, 3)?.Trim(), "Volleyball", StringComparison.OrdinalIgnoreCase)) volleyballCount++;
                    if (string.Equals(worksheet.GetText(i, 3)?.Trim(), "Basketball", StringComparison.OrdinalIgnoreCase)) basketballCount++;
                    if (string.Equals(worksheet.GetText(i, 3)?.Trim(), "Tennis", StringComparison.OrdinalIgnoreCase)) tennisCount++;
                    if (string.Equals(worksheet.GetText(i, 3)?.Trim(), "Soccer", StringComparison.OrdinalIgnoreCase)) soccerCount++;
                    if (string.Equals(worksheet.GetText(i, 3)?.Trim(), "Badminton", StringComparison.OrdinalIgnoreCase)) badmintonCount++;
                    if (string.Equals(worksheet.GetText(i, 3)?.Trim(), "Baseball", StringComparison.OrdinalIgnoreCase)) baseballCount++;

                    string course = worksheet.GetText(i, 7)?.Trim();
                    if (string.Equals(course, "BSIT", StringComparison.OrdinalIgnoreCase)) bsitCount++;
                    else if (string.Equals(course, "BSTM", StringComparison.OrdinalIgnoreCase)) bstmCount++;
                    else if (string.Equals(course, "BEED", StringComparison.OrdinalIgnoreCase)) beedCount++;

                    if (string.Equals(worksheet.GetText(i, 6)?.Trim(), "Orange", StringComparison.OrdinalIgnoreCase)) orangeCount++;
                    if (string.Equals(worksheet.GetText(i, 6)?.Trim(), "Blue", StringComparison.OrdinalIgnoreCase)) blueCount++;
                    if (string.Equals(worksheet.GetText(i, 6)?.Trim(), "Yellow", StringComparison.OrdinalIgnoreCase)) yellowCount++;
                    if (string.Equals(worksheet.GetText(i, 6)?.Trim(), "Red", StringComparison.OrdinalIgnoreCase)) redCount++;
                }

                lblActivestudentsCount.Text = activeCount.ToString();
                lblInactivestudentscount.Text = inactiveCount.ToString();
                lblMaleCount.Text = maleCount.ToString();
                lblFemaleCount.Text = femaleCount.ToString();
                lblVolleyballCount.Text = volleyballCount.ToString();
                lblBaseballCount.Text = basketballCount.ToString();
                lblTennisCount.Text = tennisCount.ToString();
                lblSoccerCount.Text = soccerCount.ToString();
                lblBaseballCount.Text = badmintonCount.ToString();
                lblBaseballCount.Text = baseballCount.ToString();
                lblBsitCount.Text = bsitCount.ToString();
                lblBstmCount.Text = bstmCount.ToString();
                lblBeedCount.Text = beedCount.ToString();
                lblOrangeCount.Text = orangeCount.ToString();
                lblBlueCount.Text = blueCount.ToString();
                lblYellowCount.Text = yellowCount.ToString();
                lblRedCount.Text = redCount.ToString();
            }
        }


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
          
        }

        private void btnMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        

        private void btnClosewindow_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnActivestudent_Click(object sender, EventArgs e)
        {
           
        }

        private void btnLogs_Click(object sender, EventArgs e)
        {
            
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

        

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            login.Show();
            this.Close();
        }

        private void btnAddstudent_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Show();
            this.Hide();
        }

}
