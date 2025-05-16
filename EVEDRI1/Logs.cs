using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EVEDRI1
{
    public partial class Logs : Form
    {
        
        
        public Logs()
        {
            InitializeComponent();
            LoadLogsFromExcel();


        }
        public void LoadLogsFromExcel()
        {
           Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx");
            Worksheet sheet = workbook.Worksheets[2];

            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();
            dgvLogs.DataSource = dt;
        }

      

        private void lblReturnDashboard_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard();
            dashboard.Show();
            this.Hide();
        }

        private void btnSearchlog_Click(object sender, EventArgs e)
        {
           
            string searchText = txtSearch.Text.ToLower();
            bool foundMatch = false;

            if (string.IsNullOrEmpty(txtSearch.Text))
            {
                MessageBox.Show("Please enter the cell you want to search.", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }



            foreach (DataGridViewRow row in dgvLogs.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().ToLower().Split(' ').Contains(searchText))
                    {
                        cell.Style.BackColor = Color.Yellow;
                        foundMatch = true;
                    }
                    else
                    {
                        cell.Style.BackColor = dgvLogs.DefaultCellStyle.BackColor;
                    }
                }
            }

            if (foundMatch)
            {
                MessageBox.Show("Matching cells have been highlighted.", "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("No matching cells found.", "Search Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text.ToLower();

            foreach (DataGridViewRow row in dgvLogs.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Style.BackColor = dgvLogs.DefaultCellStyle.BackColor;
                }
            }
        }
    }
}
