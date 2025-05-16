using Spire.Xls;
using System;
using System.Data;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace EVEDRI1
{
    public partial class Form1 : Form
    {

        string rad;
        string chk;
        Form2 frm2;
        Workbook book = new Workbook();


        public Form1()
        {
            InitializeComponent();
            frm2 = new Form2();
            dtpBirthday.ValueChanged += dtpBirthday_ValueChanged;
            txtAge.ReadOnly = true;
        }

        public bool IsValidEmail(string email)
        {

            string pattern = @"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$";
            Regex regex = new Regex(pattern);
            return regex.IsMatch(email);
        }
        public int CalculateAge(DateTime birthAge)
        {
            int age = DateTime.Today.Year - birthAge.Year;
            if (birthAge > DateTime.Today.AddYears(-age)) age--;
            return age;
        }
        private void dtpBirthday_ValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = dtpBirthday.Value;
            int age = CalculateAge(selectedDate);
            txtAge.Text = age.ToString();
        }

        public void UpdateTextFields(int ID, string name, string gender, string sports, string address, string email, string birthday, string age, string favColor, string user, string pass, string saying, string course, string status, string profile)
        {
            txtName.Text = name;

            ID = Convert.ToInt32(ID);

            if (gender == "Male")
            {
                radMale.Checked = true;
            }
            else if (gender == "Female")
            {
                radFemale.Checked = true;
            }

            chkBasketball.Checked = sports.Contains("Basketball");
            chkVolleyball.Checked = sports.Contains("Volleyball");
            chkBadminton.Checked = sports.Contains("Badminton");

            txtAddress.Text = address;
            txtEmail.Text = email;
            dtpBirthday.Text = birthday;
            txtAge.Text = age;
            cbmFavcolor.Text = favColor;
            txtUsername.Text = user;
            txtPassword.Text = pass;
            txtSaying.Text = saying;
            cbmCourse.Text = course;
            lblStatus.Text = status;
            lblProfile.Text = profile;
        }

       

        private void btnReturn_Click(object sender, EventArgs e)
        {
           Mylogs mylogs = new Mylogs();
            mylogs.insertLog(Props.CurrentUser, "Returned to the active list.");
            Form2 frm2 = new Form2();
            frm2.Show();
            this.Hide();
        }

     

        private void btnAdd_Click_1(object sender, EventArgs e)
        {
            Mylogs logs = new Mylogs();
            logs.insertLog(Props.CurrentUser, "Added a New Student.");
            lblErrors.Text = "";
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(txtName.Text)) errors.AppendLine("• Name is required.");
            if (!radMale.Checked && !radFemale.Checked) errors.AppendLine("• Gender is required.");
            if (!chkBasketball.Checked && !chkVolleyball.Checked && !chkBadminton.Checked) errors.AppendLine("• At least one sport must be selected.");
            if (string.IsNullOrWhiteSpace(txtAddress.Text)) errors.AppendLine("• Address is required.");
            if (string.IsNullOrWhiteSpace(txtEmail.Text)) errors.AppendLine("• Email is required.");
            if (!dtpBirthday.Checked) errors.AppendLine("• Birthday is required.");
            if (string.IsNullOrWhiteSpace(txtAge.Text)) errors.AppendLine("• Age is required.");
            if (cbmFavcolor.SelectedIndex == -1) errors.AppendLine("• Favorite color must be selected.");
            if (string.IsNullOrWhiteSpace(txtUsername.Text)) errors.AppendLine("• Username is required.");
            if (string.IsNullOrWhiteSpace(txtPassword.Text)) errors.AppendLine("• Password is required.");
            if (string.IsNullOrWhiteSpace(txtSaying.Text)) errors.AppendLine("• Saying is required.");
            if (cbmCourse.SelectedIndex == -1) errors.AppendLine("• Course must be selected.");
            if (string.IsNullOrWhiteSpace(txtProfilePath.Text)) errors.AppendLine("• Profile is required.");

            DateTime birthDate = dtpBirthday.Value;
            int calculatedAge = CalculateAge(birthDate);
            txtAge.Text = calculatedAge.ToString();


            if (errors.Length > 0)
            {
                lblErrors.Text = errors.ToString();
                lblErrors.Visible = true;
                MessageBox.Show("Please fill in all required fields!", "MISSING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtName.Clear();
                txtAddress.Clear();
                txtEmail.Clear();
                dtpBirthday.Checked = false;
                txtUsername.Clear();
                txtPassword.Clear();
                cbmFavcolor.SelectedIndex = -1;
                cbmCourse.SelectedIndex = -1;
                radMale.Checked = false;
                radFemale.Checked = false;
                chkBasketball.Checked = false;
                chkVolleyball.Checked = false;
                chkBadminton.Checked = false;
                chkSoccer.Checked = false;
                chkTennis.Checked = false;
                chkBaseball.Checked = false;
                txtAge.Clear();
                txtSaying.Clear();
                txtProfilePath.Clear();
                return;

            }

            try
            {

                string name = txtName.Text;
                if (radMale.Checked)
                {
                    rad = "Male";
                }
                if (radFemale.Checked)
                {
                    rad = "Female";
                }

                if (chkBasketball.Checked) chk += "Basketball ";
                if (chkVolleyball.Checked) chk += "Volleyball ";
                if (chkBadminton.Checked) chk += "Badminton ";

                string address = txtAddress.Text;
                string email = txtEmail.Text;
                string birthday = dtpBirthday.Text;
                string age = txtAge.Text;
                string favColor = cbmFavcolor.Text;
                string user = txtUsername.Text;
                string pass = txtPassword.Text;
                string saying = txtSaying.Text;
                string course = cbmCourse.Text;
                string profile = txtProfilePath.Text;
                string status = lblStatus.Text;



                if (!IsValidEmail(email))
                {
                    MessageBox.Show("Invalid email format. Please enter a valid email.");
                    return;
                }
                book.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx");
                Worksheet sheet = book.Worksheets[0];

                for (int row = 2; row <= sheet.LastRow; row++)//ERROR FOR EXISTING USER AND PASS
                {
                    string existingUsername = sheet.Range[row, 9].Value;
                    string existingPassword = sheet.Range[row, 10].Value;

                    if (existingUsername == txtUsername.Text)
                    {
                        MessageBox.Show("Username already exists. Please choose a different one.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (existingPassword == txtPassword.Text)
                    {
                        MessageBox.Show("Password already exists. Please choose a different one.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                int i = sheet.Rows.Length + 1;
                sheet.Range[i, 1].Value = name;
                sheet.Range[i, 2].Value = rad;
                sheet.Range[i, 3].Value = chk;
                sheet.Range[i, 4].Value = address;
                sheet.Range[i, 5].Value = email;
                sheet.Range[i, 6].Value = birthday;
                sheet.Range[i, 7].Value = age;
                sheet.Range[i, 8].Value = favColor;
                sheet.Range[i, 9].Value = user;
                sheet.Range[i, 10].Value = pass;
                sheet.Range[i, 11].Value = saying;
                sheet.Range[i, 12].Value = course;
                sheet.Range[i, 13].Value = status;
                sheet.Range[i, 14].Value = profile;

                book.SaveToFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx", ExcelVersion.Version2016);

                DialogResult result = MessageBox.Show("Student successfully added!", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (result == DialogResult.OK)
                {
                    Dashboard dashboard = new Dashboard();
                    Mylogs mylogs = new Mylogs();
                    mylogs.insertLog(Props.CurrentUser, "Added a new Student to the list.");
                    dashboard.Show();
                }

                txtName.Clear();
                txtAddress.Clear();
                txtEmail.Clear();
                dtpBirthday.Checked = false;
                txtUsername.Clear();
                txtPassword.Clear();
                cbmFavcolor.SelectedIndex = -1;
                cbmCourse.SelectedIndex = -1;
                radMale.Checked = false;
                radFemale.Checked = false;
                chkBasketball.Checked = false;
                chkVolleyball.Checked = false;
                chkBadminton.Checked = false;
                txtAge.Clear();
                txtSaying.Clear();
                txtProfilePath.Clear();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
           
                OpenFileDialog d = new OpenFileDialog();
                d.InitialDirectory = @"C:\Users\HF\Desktop\Main\Profiles";
                d.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";

                if (d.ShowDialog() == DialogResult.OK)
                {
                    txtProfilePath.Text = d.FileName;
                }
           
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Mylogs logs = new Mylogs();
            logs.insertLog(Props.CurrentUser, "Updated a Student.");
            lblErrors.Text = "";
            StringBuilder errors = new StringBuilder();

            if (string.IsNullOrWhiteSpace(txtName.Text)) errors.AppendLine("• Name is required.");
            if (!radMale.Checked && !radFemale.Checked) errors.AppendLine("• Gender is required.");
            if (!chkBasketball.Checked && !chkVolleyball.Checked && !chkBadminton.Checked) errors.AppendLine("• At least one sport must be selected.");
            if (string.IsNullOrWhiteSpace(txtAddress.Text)) errors.AppendLine("• Address is required.");
            if (string.IsNullOrWhiteSpace(txtEmail.Text)) errors.AppendLine("• Email is required.");
            if (!dtpBirthday.Checked) errors.AppendLine("• Birthday is required.");
            if (string.IsNullOrWhiteSpace(txtAge.Text)) errors.AppendLine("• Age is required.");
            if (cbmFavcolor.SelectedIndex == -1) errors.AppendLine("• Favorite color must be selected.");
            if (string.IsNullOrWhiteSpace(txtUsername.Text)) errors.AppendLine("• Username is required.");
            if (string.IsNullOrWhiteSpace(txtPassword.Text)) errors.AppendLine("• Password is required.");
            if (string.IsNullOrWhiteSpace(txtSaying.Text)) errors.AppendLine("• Saying is required.");
            if (cbmCourse.SelectedIndex == -1) errors.AppendLine("• Course must be selected.");


            if (errors.Length > 0)
            {
                lblErrors.Text = errors.ToString();
                lblErrors.Visible = true;
                MessageBox.Show("Please fill in all the required fields!", "MISSING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtName.Clear();
                txtAddress.Clear();
                txtEmail.Clear();
                dtpBirthday.Checked = false;
                txtUsername.Clear();
                txtPassword.Clear();
                cbmFavcolor.SelectedIndex = -1;
                cbmCourse.SelectedIndex = -1;
                radMale.Checked = false;
                radFemale.Checked = false;
                chkBasketball.Checked = false;
                chkVolleyball.Checked = false;
                chkBadminton.Checked = false;
                txtAge.Clear();
                txtSaying.Clear();
                txtProfilePath.Clear();
                return;
            }

            try
            {
                Dashboard dashboard = new Dashboard();
                Workbook book = new Workbook();
                book.LoadFromFile(@"C:\Users\HF\Desktop\Main\Book1.xlsx");
                Worksheet sheet = book.Worksheets[0];

                if (radMale.Checked)
                {
                    rad = "Male";
                }
                if (radFemale.Checked)
                {
                    rad = "Female";
                }

                if (chkBasketball.Checked) chk += "Basketball ";
                if (chkVolleyball.Checked) chk += "Volleyball ";
                if (chkBadminton.Checked) chk += "Badminton ";

                string name = txtName.Text;
                string saying = txtSaying.Text;
                string favColor = cbmFavcolor.Text;
                string address = txtAddress.Text;
                string birthday = dtpBirthday.Text;
                string age = txtAge.Text;
                string email = txtEmail.Text;
                string user = txtUsername.Text;
                string pass = txtPassword.Text;
                string course = cbmCourse.Text;
                string status = lblStatus.Text;
                string profile = lblProfile.Text;

                if (!IsValidEmail(email))
                {
                    MessageBox.Show("Invalid email format. Please enter a valid email.");
                    return;
                }
                int ID = Convert.ToInt32(lblId.Text);

                Form2 frm2 = new Form2();
                frm2.UpdateToExcel(ID, name, rad, chk, address, email, birthday, age, favColor, user, pass, saying, course, status, profile);

                DialogResult result = MessageBox.Show("Student details updated successfully!", "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (result == DialogResult.OK)
                {
                    frm2.Show();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnReturn_Click_1(object sender, EventArgs e)
        {
            Dashboard dashboard = new Dashboard();
        }

        private void btnClosewindow_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
