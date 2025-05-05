using Spire.Xls;
using System;
using System.Data;
using System.Windows.Forms;


namespace EVEDRI1
{
    public partial class Form1 : Form
    {
        string name = "";
        string gender = "";
        string hobbies = "";
        string favcolor = "";
        string saying = "";
        string age = "";
        string birthday = "";
        string email = "";
        string username = "";
        string password = "";
        string address = "";
        string course = "";



        Form2 form2 = new Form2();
        public Form1()
        {
            InitializeComponent();
            lblId.Visible = false;

        }

        public string checkEmpty()
        {
            string error = "";
            foreach (Control c in Controls)
            {
                if (c is TextBox)
                {
                    if (c.Text == "")
                    {
                        error += c.Name + "is empty";
                    }
                }
            }
            return error;
        }

        private void btnAdddata_Click(object sender, EventArgs e)
        {
            lblMessage.Text = checkEmpty();
            //Form1 Data Adding
            name = txtName.Text;
            email = txtEmail.Text;
            address = txtAddress.Text;
            if (radMale.Checked)
            {
                gender += "Male";
            }
            else
            {
                gender += "Female";
            }
            if (chkVolleyball.Checked == true)
            {
                hobbies += "Volleyball";
            }
            if (chkBasketball.Checked == true)
            {
                hobbies += "Basketball";
            }
            if (chkBadminton.Checked == true)
            {
                hobbies += "Badminton";
            }
            if (chkTennis.Checked == true)
            {
                hobbies += "Tennis";
            }
            if (chkSoccer.Checked == true)
            {
                hobbies += "Soccer";
            }
            if (chkBaseball.Checked == true)
            {
                hobbies += "Baseball";
            }
            course = cbmCourse.Text;
            favcolor = cbmFavcolor.Text;
            saying = txtSaying.Text;
            username = txtUsername.Text;
            password = txtPassword.Text;
            address = txtAddress.Text;
            birthday = dtpBirthday.Text;
            age = txtAge.Text;
            //form2.Insert(txtName.Text, gender, hobbies, favcolor, saying, saying);

            //Excel Insert
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\HF\Desktop\EVEDRI1latest\EVEDRI1\Book1.xlsx"); //Change file location
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length + 1;
            sheet.Range[row, 1].Value = name;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = birthday;
            sheet.Range[row, 5].Value = age;
            sheet.Range[row, 6].Value = favcolor;
            sheet.Range[row, 7].Value = course;
            sheet.Range[row, 8].Value = address;
            sheet.Range[row, 9].Value = saying;
            sheet.Range[row, 10].Value = email;
            sheet.Range[row, 11].Value = username;
            sheet.Range[row, 12].Value = password;


            // Save Excel
            workbook.SaveToFile(@"C:\Users\HF\Desktop\EVEDRI1latest\EVEDRI1\Book1.xlsx", ExcelVersion.Version2016); // Change file location
            DataTable dt = sheet.ExportDataTable();
            form2.dataGridView1.DataSource = dt;

            txtName.Text = string.Empty;
            txtSaying.Text = string.Empty;
            radFemale.Checked = false;
            radMale.Checked = false;
            chkVolleyball.Checked = false;
            chkBasketball.Checked = false;
            chkBadminton.Checked = false;
            chkTennis.Checked = false;
            chkSoccer.Checked = false;
            chkBaseball.Checked = false;
            cbmFavcolor.Text = string.Empty;
            txtSaying.Text = string.Empty;
            txtUsername.Text = string.Empty;
            txtPassword.Text = string.Empty;
            txtEmail.Text = string.Empty;

            dtpBirthday.Text = string.Empty;
            txtAge.Text = string.Empty;
            txtAddress.Text = string.Empty;
            cbmCourse.Text = string.Empty;
            txtUsername.Text = string.Empty;
            txtPassword.Text = string.Empty;
            dtpBirthday.Text = string.Empty;
            txtAge.Text = string.Empty;

            cbmCourse.SelectedIndex = -1;
            cbmFavcolor.SelectedIndex = -1;
            dtpBirthday.CustomFormat = string.Empty;
        }

        private void btnDisplayall_Click(object sender, EventArgs e)
        {
            this.Hide();
            form2.Show();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            //Data
            name = txtName.Text;
            if (radMale.Checked)
            {
                gender += "Male";
            }
            else
            {
                gender += "Female";
            }
            if (chkVolleyball.Checked == true)
            {
                hobbies += "Volleyball";
            }
            if (chkBasketball.Checked == true)
            {
                hobbies += "Basketball";
            }
            if (chkBadminton.Checked == true)
            {
                hobbies += "Badminton";
            }
            if (chkTennis.Checked == true)
            {
                hobbies += "Tennis";
            }
            if (chkSoccer.Checked == true)
            {
                hobbies += "Soccer";
            }
            if (chkBaseball.Checked == true)
            {
                hobbies += "Baseball";
            }
            favcolor = cbmFavcolor.Text;
            saying = txtSaying.Text;
            username = txtUsername.Text;
            password = txtPassword.Text;

            //form2.Update(Convert.ToInt32(lblId.Text),txtName.Text, gender, favcolor, hobbies, txtSaying.Text);

            //Excel Update
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\HF\Desktop\EVEDRI1latest\EVEDRI1\Book1.xlsx"); //Change file location
            Worksheet sheet = workbook.Worksheets[0];
            int row = form2.dataGridView1.CurrentCell.RowIndex + 3;
            sheet.Range[row, 1].Value = name;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = birthday;
            sheet.Range[row, 5].Value = age;
            sheet.Range[row, 6].Value = favcolor;
            sheet.Range[row, 7].Value = course;
            sheet.Range[row, 8].Value = address;
            sheet.Range[row, 9].Value = saying;
            sheet.Range[row, 10].Value = email;
            sheet.Range[row, 11].Value = username;
            sheet.Range[row, 12].Value = password;

            //Excel Save 
            workbook.SaveToFile(@"C:\Users\HF\Desktop\EVEDRI1latest\EVEDRI1\Book1.xlsx", ExcelVersion.Version2016); //Change file location
            DataTable dt = sheet.ExportDataTable();
            form2.dataGridView1.DataSource = dt;
            form2.Show();
            //int row;
        }

        private void dtpBirthday_ValueChanged(object sender, EventArgs e)
        {
            CalcuAge();
        }

        private void CalcuAge()
        {
            DateTime birthday = dtpBirthday.Value.Date;
            DateTime today = DateTime.Today;

            int age = today.Year - birthday.Year;

            if (birthday > today.AddYears(-age))
            {
                age--;
            }
            txtAge.Text = age.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnClosewindow_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDashboard_Click(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard();
            dashboard.Show();
            this.Hide();
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login login = new Login();
            login.Show();
            this.Close();
        }
    }
}
