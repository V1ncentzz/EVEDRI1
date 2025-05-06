using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management.Instrumentation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace EVEDRI1
{
    public partial class Form2 : Form
    {
        
        Workbook book = new Workbook();
        public Form2()
        {
            InitializeComponent();
           
            LoadExcelFile();
            ShowStudents("Status =");
        }
        
        public void LoadExcelFile()
        {
            //book.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx"); //Change file location
            //Worksheet sheet = book.Worksheets[0];
            //DataTable dt = sheet.ExportDataTable();
            //dataGridView1.DataSource = dt;

            book.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable fullTable = sheet.ExportDataTable();
            DataTable activeStudents = fullTable.Clone();

            foreach (DataRow row in fullTable.Rows)
            {
                if (row[12].ToString() == "1") // 13th column = Status
                {
                    activeStudents.ImportRow(row);
                }
            }

            dataGridView1.DataSource = activeStudents;


        }
        
        public void ShowStudents(string status)
        {
            book.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx");
            Worksheet worksheet = book.Worksheets[0];
            DataTable dt = worksheet.ExportDataTable();
            //DataRow[] rows = dt.Select("Status=" + status);

            //foreach (DataRow i in rows)
            //{
            //    dataGridView1.Rows.Insert([0]
            //}
        }



        public void Insert(string name, string gender, string hobbies, string birthday, string age, string favcolor, 
        string course, string address, string saying, string email, 
        string username, string password)
        {
            int row = dataGridView1.Rows.Add();
            dataGridView1.Rows[row].Cells[0].Value = name;
            dataGridView1.Rows[row].Cells[1].Value = gender;
            dataGridView1.Rows[row].Cells[2].Value = hobbies;
            dataGridView1.Rows[row].Cells[3].Value = birthday;
            dataGridView1.Rows[row].Cells[4].Value = age;
            dataGridView1.Rows[row].Cells[5].Value = favcolor;
            dataGridView1.Rows[row].Cells[6].Value = course;
            dataGridView1.Rows[row].Cells[7].Value = address;
            dataGridView1.Rows[row].Cells[8].Value = saying;
            dataGridView1.Rows[row].Cells[9].Value = email;
            dataGridView1.Rows[row].Cells[10].Value = username;
            dataGridView1.Rows[row].Cells[11].Value = password;
        }

        public void Update(int id, string name, string gender, string hobbies, string birthday, string age, string favcolor, string course, 
        string address, string saying, string email, 
        string username, string password)
        {
            //int i = dataGridView1.Rows.Add();
            dataGridView1.Rows[id].Cells[0].Value = name;
            dataGridView1.Rows[id].Cells[1].Value = gender;
            dataGridView1.Rows[id].Cells[2].Value = hobbies;
            dataGridView1.Rows[id].Cells[3].Value = birthday;
            dataGridView1.Rows[id].Cells[4].Value = age;
            dataGridView1.Rows[id].Cells[5].Value = favcolor;
            dataGridView1.Rows[id].Cells[6].Value = course;
            dataGridView1.Rows[id].Cells[7].Value = address;
            dataGridView1.Rows[id].Cells[8].Value = saying;
            dataGridView1.Rows[id].Cells[9].Value = email;
            dataGridView1.Rows[id].Cells[10].Value = username;
            dataGridView1.Rows[id].Cells[11].Value = password;
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchVal = txtSearch.Text;
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            string filterVal = searchVal.Replace("'", "''");
            bs.Filter = $"Name LIKE '{filterVal}%'";
            dataGridView1.DataSource = bs;

           
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form1 form1 = (Form1)Application.OpenForms["Form1"];
            int r = dataGridView1.CurrentCell.RowIndex;
            form1.txtName.Text = dataGridView1.Rows[r].Cells[0].Value.ToString();
            form1.cbmCourse.Text = dataGridView1.Rows[r].Cells[6].Value.ToString();
            form1.txtEmail.Text = dataGridView1.Rows[r].Cells[9].Value.ToString();
            form1.txtAddress.Text = dataGridView1.Rows[r].Cells[7].Value.ToString();
            form1.txtSaying.Text = dataGridView1.Rows[r].Cells[4].Value.ToString();
            form1.cbmFavcolor.Text = dataGridView1.Rows[r].Cells[3].Value.ToString();
            form1.txtUsername.Text = dataGridView1.Rows[r].Cells[10].Value.ToString();
            form1.txtPassword.Text = dataGridView1.Rows[r].Cells[11].Value.ToString();
            
            form1.lblId.Text = r.ToString();
            switch (dataGridView1.Rows[r].Cells[1].Value)
            {
                case "Male":
                    form1.radMale.Checked = true;
                    break;
                default:
                    form1.radFemale.Checked = true;
                    break;
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
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

        //private void btnDelete_Click(object sender, EventArgs e)
        //{
        //    int selectedRow = dataGridView1.CurrentCell.RowIndex;
        //    Worksheet sheet = book.Worksheets[0];
        //    sheet.Range[selectedRow + 2, 13].Value = "0"; // Set status to inactive
        //    book.SaveToFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx", ExcelVersion.Version2016);
        //    LoadActiveStudents();

        //    //Worksheet sheet = book.Worksheets[0];
        //    //DataTable dt = sheet.ExportDataTable();
        //    //dataGridView1.DataSource = dt;
        //    //int row = sheet.Rows.Length;
        //    //sheet.DeleteRow(row);
        //    //foreach (DataGridViewRow row in dataGridView1.SelectedRows)
        //    //{
        //    //    dataGridView1.Rows.Remove(row);
        //    //}
        //}

        private void LoadActiveStudents()
        {
            book.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            DataView dv = dt.DefaultView;
            dv.RowFilter = "F13 = '1'"; // Assuming F13 is "Status"
            dataGridView1.DataSource = dv.ToTable();
        }
        


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"C:\Users\\ACT-STUDENT\\Desktop\\Book1");
            //Worksheet sheet = book.Worksheets[0];
            //int row = sheet.Rows.Length;
            //for (int i = 0; i <)
        }

        private void lblSearch_Click(object sender, EventArgs e)
        {

        }

        //private void Form2_Load(object sender, EventArgs e)
        //{

        //}

        private void btnClosewindow_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDeteteRow_Click(object sender, EventArgs e)
        {
            int selectedRowIndex = dataGridView1.CurrentCell.RowIndex;

            // Load Excel
            book.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx");
            Worksheet sheet = book.Worksheets[0];

            // Update the status to 0
            sheet.Range[selectedRowIndex + 2, 13].Value = "0"; // Set Status to Inactive (row + 2 for Excel index)

            book.SaveToFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx", ExcelVersion.Version2016);

            MessageBox.Show("Student marked as inactive.");
            LoadExcelFile(); // Refresh

            //    Worksheet sheet = book.Worksheets[0];
            //    DataTable dt = sheet.ExportDataTable();
            //    dataGridView1.DataSource = dt;
            //    int row = sheet.Rows.Length;
            //    sheet.DeleteRow(row);
            //    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            //    {
            //        dataGridView1.Rows.Remove(row);
            //    }
        }

            private void btnInactive_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3();
            form3.Show();
            this.Close();
        }
    }
}
