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
using static Guna.UI2.Native.WinApi;

namespace EVEDRI1
{
    public partial class Form3 : Form
    {
        Workbook workbook = new Workbook();

        public Form3()
        {
            InitializeComponent();
            LoadInactiveStudents();
        }

        public void LoadExcelFile()
        {
            workbook.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx"); //Change file location
            Worksheet sheet = workbook.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            DataGridInactive.DataSource = dt;
        }

        private void LoadInactiveStudents()
        {
            workbook.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            DataTable fullTable = sheet.ExportDataTable();
            DataTable inactive = fullTable.Clone();

            foreach (DataRow row in fullTable.Rows)
            {
                if (row[12].ToString() == "0") // Status = Inactive
                {
                    inactive.ImportRow(row);
                }
            }

            DataGridInactive.DataSource = inactive;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form1 form1 = (Form1)Application.OpenForms["Form1"];
            int r = DataGridInactive.CurrentCell.RowIndex;
            form1.txtName.Text = DataGridInactive.Rows[r].Cells[0].Value.ToString();
            form1.cbmCourse.Text = DataGridInactive.Rows[r].Cells[6].Value.ToString();
            form1.txtEmail.Text = DataGridInactive.Rows[r].Cells[9].Value.ToString();
            form1.txtAddress.Text = DataGridInactive.Rows[r].Cells[7].Value.ToString();
            form1.txtSaying.Text = DataGridInactive.Rows[r].Cells[4].Value.ToString();
            form1.cbmFavcolor.Text = DataGridInactive.Rows[r].Cells[3].Value.ToString();
            form1.txtUsername.Text = DataGridInactive.Rows[r].Cells[10].Value.ToString();
            form1.txtPassword.Text = DataGridInactive.Rows[r].Cells[11].Value.ToString();

            form1.lblId.Text = r.ToString();
            switch (DataGridInactive.Rows[r].Cells[1].Value)
            {
                case "Male":
                    form1.radMale.Checked = true;
                    break;
                default:
                    form1.radFemale.Checked = true;
                    break;
            }

            foreach (DataGridViewRow row in DataGridInactive.Rows)
            {
                if (!row.IsNewRow)
                {
                    string cellValue = row.Cells[2].Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        string[] lines = cellValue.Split(' ');

                        foreach (string line in lines)
                        {
                            if (line == "Volleyball")
                            {
                                form1.chkVolleyball.Checked = true;
                            }
                            if (line == "Basketball")
                            {
                                form1.chkBasketball.Checked = true;
                            }
                            if (line == "Badminton")
                            {
                                form1.chkBadminton.Checked = true;
                            }
                            if (line == "Tennis")
                            {
                                form1.chkTennis.Checked = true;
                            }
                            if (line == "Soccer")
                            {
                                form1.chkSoccer.Checked = true;
                            }
                            if (line == "Baseball")
                            {
                                form1.chkBaseball.Checked = true;
                            }


                        }
                    }
                }
            }
            this.Hide();
            form1.Show();
            form1.btnUpdate.Visible = true;
            form1.btnDisplayall.Visible = false;
            form1.btnAdddata.Visible = false;
        }

        private void btnReactivate_Click(object sender, EventArgs e)
        {
            int rowIndex = DataGridInactive.CurrentCell.RowIndex;
            string studentName = DataGridInactive.Rows[rowIndex].Cells[0].Value.ToString();

            workbook.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            int totalRows = sheet.Rows.Length;

            for (int i = 2; i <= totalRows; i++) // Start at 2 to skip headers
            {
                if (sheet.Range[i, 1].Value == studentName)
                {
                    sheet.Range[i, 13].Value = "1"; // Set Status to Active
                    break;
                }
            }

            workbook.SaveToFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx", ExcelVersion.Version2016);
            MessageBox.Show("Student reactivated.");

            LoadInactiveStudents();
        }
    }
}
