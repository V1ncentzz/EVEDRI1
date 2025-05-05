using Guna.UI2.WinForms;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace EVEDRI1
{
    public partial class Login : Form
    {
        Dashboard d = new Dashboard();
        public Login()
        {
            InitializeComponent();
            txtUsername.Text = "Username";
            txtUsername.ForeColor = SystemColors.GrayText;
            txtPassword.Text = "Password";
            txtPassword.ForeColor = SystemColors.GrayText;

            txtUsername.Enter += txtUsername_Enter;
            txtUsername.Leave += txtUsername_Leave;
            txtPassword.Enter += txtPassword_Enter;
            txtPassword.Leave += txtPassword_Leave;

            this.MouseDown += Login_MouseDown;
            // Hide border (important!)
            this.FormBorderStyle = FormBorderStyle.None;
            // Center the form
            this.StartPosition = FormStartPosition.CenterScreen;
            ///panel1.MouseDown += Login_MouseDown;
            //guna2GradientPanel1.MouseDown += Login_MouseDown;

            Mylogs mylogs = new Mylogs();
            mylogs.insertlog("user", "message");
           
        }

        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        private void Login_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, 0xA1, 0x2, 0);
            }
        }



        private void txtUsername_Enter(object sender, EventArgs e)
        {
            if (txtUsername.Text == "Username")
            {
                txtUsername.Text = "";
                txtUsername.ForeColor = SystemColors.WindowText;
            }
        }

        private void txtUsername_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text))
            {
                txtUsername.Text = "Username";
                txtUsername.ForeColor = SystemColors.GrayText;
            }
        }

        private void txtPassword_Enter(object sender, EventArgs e)
        {
            if (txtPassword.Text == "Password")
            {
                txtPassword.Text = "";
                txtPassword.ForeColor = SystemColors.WindowText;
            }
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                txtPassword.Text = "Password";
                txtPassword.ForeColor = SystemColors.GrayText;
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
            //Textbox Username
            txtUsername.IconLeft = Image.FromFile("ICON USERNAME.png");
            txtUsername.IconLeftSize = new Size(40,40);
            txtUsername.ShadowDecoration.Enabled = true;
            txtUsername.ShadowDecoration.Color = Color.Gray;
            txtUsername.ShadowDecoration.Depth = 5;

            //Textbox Password
            txtPassword.IconLeft = Image.FromFile("ICON PASSWORD.png");
            txtPassword.IconLeftSize = new Size(40, 40);
            txtPassword.ShadowDecoration.Enabled = true;
            txtPassword.ShadowDecoration.Color = Color.Gray;
            txtPassword.ShadowDecoration.Depth = 5;

            //Butto Login
            btnLogin.BorderRadius = 10;

        }



        private void btnLogin_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\HF\Documents\Ff\Book1.xlsx"); //Change file location
            Worksheet sheet = workbook.Worksheets[0];
            int row = sheet.Rows.Length;
            bool login = false;

            for (int i = 2; i <= row; i++)
            {
                if (sheet.Range[i,11].Value == txtUsername.Text && sheet.Range[i, 12].Value == txtPassword.Text)
                {
                    d.picUser.Image = Image.FromFile(@"C:\Users\HF\Desktop\EVEDRI1latest\EVEDRI1\EVEDRI1\Resources\profileicon.jpg" + sheet.Range[i,14].Value);
                    d.picUser.SizeMode = PictureBoxSizeMode.StretchImage;
                    login = true;
                    break;
                }
                else
                {
                    login = false;

                }
            }

            if (login == true)
            {
                /*Form1 form = new Form1();
                form.Show();*/
                Dashboard dashboard = new Dashboard();
                dashboard.Show();
                this.Hide();

            }
            else
            {
                MessageBox.Show("Your entered incorrect username or password, please try again!");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private Rectangle prevBounds;
        private bool isMaximized = false;
        private void btnMaximize_Click(object sender, EventArgs e)
        {
            //this.WindowState = FormWindowState.Maximized;
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

        private void btnMinimize_Click(object sender, EventArgs e)
        {
           this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2GradientPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblSignin_Click(object sender, EventArgs e)
        {

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
